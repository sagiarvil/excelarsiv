#!/usr/bin/env python3
"""Hukuk Bürosu Dosya & Vekalet Komutası — A3+A8 motoru (manda v6 · Hat B Y-19).

Dosya durum makinesi, vekalet ücreti borçlandırma, tahsilat eşleştirme, süre alarmı.
Domain terimleri korunur: dosya / vekalet / tahsilat (rename yasak).
Avukatlık Kanunu / AAÜT / HMK süresi AYARLAR'dan; dış API yok.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
    KOYU_LACIVERT,
    SIFRE,
    TARİH,
    TL,
    YÜZDE,
    ad_ekle,
    baski_hazirla,
    baslik_satiri,
    dogrulama,
    genislik,
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
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)
from ortak.mutabakat_motoru import (  # noqa: E402
    dort_kova_butunluk,
    kok_neden_onerisi,
)

URUN_AD = "Hukuk Bürosu Dosya & Vekalet Komutası"
SURUM = "1.0.0"
RENK = "1A365D"
KAPASITE = 1000
DEMO_MUVEKKIL = 12
DEMO_DONEM = 24
DEMO_TAHSIL = 20
DEMO_AKIS = 18
DEMO_GIDER = 10
DEMO_KARAR = 6
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "ACIK", "durum_adi": "Açık", "sira": 1,
     "izinli_sonraki_durumlar": "VEKALET|IPTAL", "kategori": "acik"},
    {"durum_id": "VEKALET", "durum_adi": "Vekalet", "sira": 2,
     "izinli_sonraki_durumlar": "DAVADA|IPTAL|ACIK", "kategori": "acik"},
    {"durum_id": "DAVADA", "durum_adi": "Davada", "sira": 3,
     "izinli_sonraki_durumlar": "DURUSMA|IPTAL", "kategori": "acik"},
    {"durum_id": "DURUSMA", "durum_adi": "Duruşma", "sira": 4,
     "izinli_sonraki_durumlar": "SONUC|DAVADA", "kategori": "acik"},
    {"durum_id": "SONUC", "durum_adi": "Sonuç", "sira": 5,
     "izinli_sonraki_durumlar": "KAPALI|IPTAL", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 6,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 7,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

MUVEKKILLER = [
    ("MUV-001", "MV-1001", "Ayşe Yılmaz", 1.0, "EVET"),
    ("MUV-002", "MV-1002", "Mehmet Demir", 1.0, "EVET"),
    ("MUV-003", "MV-1003", "Zeynep Kaya", 1.0, "EVET"),
    ("MUV-004", "MV-1004", "Ali Çelik", 1.5, "EVET"),
    ("MUV-005", "MV-1005", "Fatma Şahin", 1.0, "EVET"),
    ("MUV-006", "MV-1006", "Hasan Aydın", 1.0, "HAYIR"),
    ("MUV-007", "MV-1007", "Elif Arslan", 2.0, "EVET"),
    ("MUV-008", "MV-1008", "Mustafa Öztürk", 1.0, "EVET"),
    ("MUV-009", "MV-1009", "Selin Koç", 1.0, "EVET"),
    ("MUV-010", "MV-1010", "Emre Yıldız", 1.0, "EVET"),
    ("MUV-011", "MV-1011", "Deniz Acar", 1.0, "EVET"),
    ("MUV-012", "MV-1012", "Burak Polat", 1.0, "EVET"),
]

# Vekalet borç satırları: id, müvekkil, dönem, tutar, vade
VEKALET_SATIR = []
for i, (mid, *_rest) in enumerate(MUVEKKILLER[:10], 1):
    VEKALET_SATIR.append((f"VEK-{i:03d}", mid, "2026-07", 350.0 + (i % 3) * 50, date(2026, 7, 10)))
for i, (mid, *_rest) in enumerate(MUVEKKILLER[:8], 11):
    VEKALET_SATIR.append((f"VEK-{i:03d}", mid, "2026-08", 350.0 + (i % 4) * 40, date(2026, 8, 10)))
# kısa/uzun no senaryosu
VEKALET_SATIR[2] = ("VEK-003", "MUV-003", "2026-07", 400.0, date(2026, 7, 10))
VEKALET_SATIR.append(("VEK-025", "MUV-011", "2026-08", 380.0, date(2026, 8, 10)))
VEKALET_SATIR.append(("VEK-026", "MUV-012", "2026-08", 360.0, date(2026, 8, 10)))

# Dosya kartları: id, müvekkil, konu, son süre, durum
DOSYA_SATIR = [
    (f"DFY-{i:03d}", MUVEKKILLER[(i - 1) % len(MUVEKKILLER)][0],
     konu, date(2026, 8, 5) + timedelta(days=i), durum)
    for i, (konu, durum) in enumerate([
        ("İşe iade davası", "DAVADA"),
        ("Alacak takibi", "VEKALET"),
        ("Boşanma dosyası", "DURUSMA"),
        ("Kira tespiti", "ACIK"),
        ("İcra takip", "SONUC"),
        ("Tazminat davası", "DAVADA"),
        ("Şirket uyuşmazlığı", "VEKALET"),
        ("Marka ihlali", "DURUSMA"),
        ("İşe iade (2)", "KAPALI"),
        ("Miras paylaşımı", "ACIK"),
        ("Kıdem ihbar", "DAVADA"),
        ("Sözleşme feshi", "IPTAL"),
    ], 1)
]

TAHSIL_SATIR = []
# eşleşenler
for i, (bid, mid, don, tut, _v) in enumerate(VEKALET_SATIR[:14], 1):
    tut_odeme = tut if i not in (5, 9) else tut - 50  # tutar farkı
    if i == 7:
        tut_odeme = tut  # tolerans içinde (0)
    TAHSIL_SATIR.append((f"ODM-{i:03d}", mid, don, tut_odeme, date(2026, 7, 15) + timedelta(days=i)))
# eşleşmeyen ödeme
TAHSIL_SATIR.append(("ODM-021", "MUV-006", "2026-06", 340.0, date(2026, 6, 20)))
TAHSIL_SATIR.append(("ODM-022", "MUV-999", "2026-08", 2500.0, date(2026, 8, 5)))  # yetim

GIDER_SATIR = [
    ("GDR-01", "Büro kira", 8000.0, 8000.0, "2026-08"),
    ("GDR-02", "Tebligat/masraf", 6500.0, 7200.0, "2026-08"),
    ("GDR-03", "Kırtasiye/dosya", 1200.0, 980.0, "2026-08"),
    ("GDR-04", "UYAP/internet", 2400.0, 2650.0, "2026-08"),
    ("GDR-05", "Bilirkişi avansı", 3500.0, 3500.0, "2026-08"),
    ("GDR-06", "Arabuluculuk ücreti", 4200.0, 3900.0, "2026-08"),
    ("GDR-07", "Mesleki sigorta", 1800.0, 1800.0, "2026-08"),
    ("GDR-08", "Muhasebe hizmeti", 2500.0, 2500.0, "2026-08"),
    ("GDR-09", "Duruşma yolu gideri", 1500.0, 2100.0, "2026-08"),
    ("GDR-10", "Harç avansı", 2000.0, 0.0, "2026-08"),
]

KARAR_SATIR = [
    ("KR-01", date(2026, 1, 15), "2026 vekalet tarife tablosu", "ONAY", "BURO"),
    ("KR-02", date(2026, 3, 20), "Müvekkil vekalet onay süreci", "ONAY", "MUVEKKIL"),
    ("KR-03", date(2026, 5, 10), "Asgari vekalet ücreti AAÜT", "ONAY", "MUVEKKIL"),
    ("KR-04", date(2026, 6, 12), "Büro arşiv dijitalleştirme", "ERTELE", "BURO"),
    ("KR-05", date(2026, 7, 8), "Gecikme zammı / süre eşiği", "ONAY", "BURO"),
    ("KR-06", date(2026, 8, 5), "Aylık dosya kapanış gündemi", "TASLAK", "BURO"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Dosya/vekalet/tahsilat kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı vekalet tahsilatı üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Hukuk bürosu dosya/vekalet kararını etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: vekalet tahsilatı/dosya akışı bozulur | Kaynak: Büro / AYARLAR",
    )


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    h(ws, 3, 1, URUN_AD, kalin=True, boyut=18, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      "Aidat borçlandırma ve tahsilatı eşleştirin; icap kuyruğunu yönetin; "
      "genel kurula sunulacak dönem kanıt raporunu üretin (AAÜT/HMK).",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:N4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Vekalet↔tahsilat dört kova mutabakatı — sapmada kök neden + TANIMSIZ kova",
        "Dosya durum makinesi: Açık → Vekalet → Davada → Duruşma → Sonuç → Kapalı",
        "AAÜT/HMK gecikme ve süre eşiği AYARLAR'dan; gömülü sabit yok",
        "Büro gideri vs bütçe + dosya/vekalet karar defteri",
        "Makrosuz, çevrimdışı; büro yazılımı / API bağı yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Avukat / dosya asistanı — dosya, vekalet ve tahsilat girişi",
        "Büro sahibi / yönetici ortak — PANO karar özeti",
        "Müvekkil / SMMM — RAPOR kanıt paketi",
    ], 14):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, f"Sürüm {SURUM} · Şifre (korumalı alanlar): {SIFRE}", yazi=GRİ)
    h(ws, 19, 1, "Başlamak: HIZLI_BASLANGIC → GIRDI → PAYDASLAR → DOSYA → VEKALET → TAHSILAT → PANO",
      kalin=True, yazi=KOYU_LACIVERT)
    genislik(ws, {"A": 88})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "3 adımda başlayın", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    for i, (ad, acik) in enumerate([
        ("1. Müvekkil & dosya", "PAYDASLAR ve DOSYA sayfalarında müvekkil/dosya kimliklerini girin (sarı hücreler)."),
        ("2. Vekalet & tahsilat", "VEKALET'te vekalet ücreti borçlandırın; TAHSILAT'ta tahsilatları girin."),
        ("3. Karar", "PANO ve KARAR'da tahsilat oranı + süre/dosya kuyruğunu okuyun."),
    ], 5):
        h(ws, i, 1, ad, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, acik, kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=10)
    h(ws, 9, 1, "Anlık karar kapısı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 1, "Karar")
    h(ws, 10, 2, "=hbd_kararKapi", kalin=True, boyut=14)
    h(ws, 11, 1, "Tahsilat oranı")
    h(ws, 11, 2, "=hbd_tahsilatOrani", sayi=YÜZDE)
    h(ws, 12, 1, "Kasa bakiyesi")
    h(ws, 12, 2, "=hbd_kasaBakiye", sayi=TL)
    genislik(ws, {"A": 22, "B": 70})


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Büro ve dönem parametreleri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    alanlar = [
        (5, "Kurum adı", "Yılmaz & Demir Hukuk Bürosu", "Metin", "Hukuk bürosu unvanı"),
        (6, "Açık dosya sayısı", 86, "Sayi", "Açık dosya adedi"),
        (7, "Dönem kodu", "2026-08", "Metin", "YYYY-AA"),
        (8, "Asgari vekalet ücreti (TL)", 350, "Para", "vekalet katsayısına çarpılır"),
        (9, "Hedef tahsilat oranı", 0.92, "Yuzde", "Altında DİKKAT/KRİTİK"),
        (10, "Süre alarm eşiği (gün)", 30, "Sayi", "Son süre sonrası gün"),
        (11, "Bütçe ayı gider tavanı", 45000, "Para", "Gider alarm eşiği"),
        (12, "Senaryo çarpanı", 1.0, "Oran", "İyimser/kötümser"),
    ]
    h(ws, 4, 1, "Alan", kalin=True)
    h(ws, 4, 2, "Değer", kalin=True)
    for r, ad, deger, tip, mes in alanlar:
        h(ws, r, 1, ad)
        sayi = TL if tip == "Para" else (YÜZDE if tip == "Yuzde" else ("0.00" if tip == "Oran" else "0"))
        _sari(ws, r, 2, deger, sayi=sayi, baslik=ad, mesaj=mes)
        if tip == "Para":
            dogrulama(ws, "decimal", "0", f"B{r}", baslik=ad, mesaj=mes,
                      hata_baslik=ad, hata_mesaj="0 veya üzeri",
                      isaret="greaterThanOrEqual")
        elif tip == "Yuzde":
            dogrulama(ws, "decimal", "0", f"B{r}", baslik=ad, mesaj="0–1 arası",
                      hata_baslik=ad, hata_mesaj="0–1", isaret="between", f2="1")
        elif tip == "Sayi":
            dogrulama(ws, "whole", "0", f"B{r}", baslik=ad, mesaj=mes,
                      hata_baslik=ad, hata_mesaj="Pozitif sayı",
                      isaret="greaterThanOrEqual")
        else:
            dogrulama(ws, "textLength", "1", f"B{r}", baslik=ad, mesaj=mes,
                      hata_baslik=ad, hata_mesaj="Boş olamaz",
                      isaret="greaterThanOrEqual", f2="80")
    genislik(ws, {"A": 28, "B": 36})


def paydaslar(ws):
    sayfa_hazirla(ws, "PAYDASLAR", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Müvekkil kartları — ID disiplini (bakiye türetilir; vekalet katsayısı)", kalin=True,
      yazi=KOYU_LACIVERT)
    sut = ["MuvekkilId", "MuvekkilNo", "AdSoyad", "VekaletKatsayi", "Aktif",
           "BorcToplam", "TahsilToplam", "Bakiye", "KayitDolu", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "BorcToplam": (
            'IF(tblPaydas[[#This Row],[MuvekkilId]]="","",'
            'SUMIF(tblVekalet[MuvekkilId],tblPaydas[[#This Row],[MuvekkilId]],tblVekalet[BorcTutari]))'
        ),
        "TahsilToplam": (
            'IF(tblPaydas[[#This Row],[MuvekkilId]]="","",'
            'SUMIF(tblTahsilat[MuvekkilId],tblPaydas[[#This Row],[MuvekkilId]],tblTahsilat[OdemeTutari]))'
        ),
        "Bakiye": (
            'IF(tblPaydas[[#This Row],[MuvekkilId]]="","",'
            'tblPaydas[[#This Row],[BorcToplam]]-tblPaydas[[#This Row],[TahsilToplam]])'
        ),
        "KayitDolu": 'IF(tblPaydas[[#This Row],[MuvekkilId]]="","",1)',
        "MotorAdim": (
            'IF(tblPaydas[[#This Row],[MuvekkilId]]="","",'
            '(tblPaydas[[#This Row],[Bakiye]]>0)+(tblPaydas[[#This Row],[Aktif]]="EVET"))'
        ),
    }
    for i, (mid, bolum, ad, pay, aktif) in enumerate(MUVEKKILLER):
        r = ILK + i
        _sari(ws, r, 1, mid, baslik="Müvekkil ID", mesaj="Benzersiz müvekkil kimliği")
        _sari(ws, r, 2, bolum, baslik="Müvekkil no", mesaj="Müvekkil numarası")
        _sari(ws, r, 3, ad, baslik="Ad soyad", mesaj="Müvekkil adı")
        _sari(ws, r, 4, pay, sayi="0.00", baslik="Vekalet katsayısı", mesaj="Standart=1; karmaşık dosya >1")
        _sari(ws, r, 5, aktif, baslik="Aktif", mesaj="EVET/HAYIR")
    tablo_ekle(ws, "tblPaydas", f"A{HDR}:J{SON}", sut, formuller=form)
    dogrulama(ws, "list", "ListeEvetHayir", f"E{ILK}:E{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Katsayı", mesaj="Katsayı ≥ 0",
              hata_baslik="Katsayı", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})
    ws.column_dimensions["C"].width = 18


def dosya(ws):
    sayfa_hazirla(ws, "DOSYA", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Dosya kartları — durum + son süre (A3 komuta)", kalin=True, yazi=KOYU_LACIVERT)
    sut = ["DosyaId", "MuvekkilId", "Konu", "SonSureTarihi", "Durum",
           "SureGun", "SureAlarm", "KayitDolu", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    yas_f = yaslandirma_gun("tblDosya[[#This Row],[SonSureTarihi]]", "raporTarihi").lstrip("=")
    form = {
        "SureGun": f'IF(tblDosya[[#This Row],[DosyaId]]="","",{yas_f})',
        "SureAlarm": (
            'IF(tblDosya[[#This Row],[DosyaId]]="","",'
            'IF(AND(tblDosya[[#This Row],[SureGun]]>hbd_sureEsikGun,'
            'OR(tblDosya[[#This Row],[Durum]]="ACIK",'
            'tblDosya[[#This Row],[Durum]]="VEKALET",'
            'tblDosya[[#This Row],[Durum]]="DAVADA",'
            'tblDosya[[#This Row],[Durum]]="DURUSMA")),"ALARM","OK"))'
        ),
        "KayitDolu": 'IF(tblDosya[[#This Row],[DosyaId]]="","",1)',
        "MotorAdim": (
            'IF(tblDosya[[#This Row],[DosyaId]]="","",'
            '(tblDosya[[#This Row],[SureAlarm]]="ALARM")+1)'
        ),
    }
    for i, (did, mid, konu, sure, durum) in enumerate(DOSYA_SATIR):
        r = ILK + i
        _sari(ws, r, 1, did, baslik="Dosya ID", mesaj="Benzersiz dosya kimliği")
        _sari(ws, r, 2, mid, baslik="Müvekkil ID", mesaj="PAYDASLAR referansı")
        _sari(ws, r, 3, konu, baslik="Konu", mesaj="Dosya konusu")
        _sari(ws, r, 4, sure, sayi=TARİH, baslik="Son süre", mesaj="HMK süre tarihi")
        _sari(ws, r, 5, durum, baslik="Durum", mesaj="Dosya durum_id")
    tablo_ekle(ws, "tblDosya", f"A{HDR}:I{SON}", sut, formuller=form)
    dogrulama(ws, "list", "ListeDurum", f"E{ILK}:E{SON}",
              baslik="Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "date", "2020-01-01", f"D{ILK}:D{SON}",
              baslik="Son süre", mesaj="Geçerli tarih",
              hata_baslik="Tarih", hata_mesaj="Tarih girin",
              isaret="greaterThanOrEqual", f2="2035-12-31")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["C"].width = 22


def vekalet(ws):
    sayfa_hazirla(ws, "VEKALET", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Vekalet ücreti borçlandırma — AAÜT tabanı + gecikme zammı", kalin=True,
      yazi=KOYU_LACIVERT)
    sut = ["BorcId", "MuvekkilId", "Donem", "BorcTutari", "VadeTarihi",
           "HamAnahtar", "GecikmeGun", "ZamTutari", "ToplamBorc", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    yas_f = yaslandirma_gun("tblVekalet[[#This Row],[VadeTarihi]]", "raporTarihi").lstrip("=")
    form = {
        "HamAnahtar": (
            'IF(tblVekalet[[#This Row],[MuvekkilId]]="","",'
            'UPPER(TRIM(SUBSTITUTE(tblVekalet[[#This Row],[MuvekkilId]]&'
            'tblVekalet[[#This Row],[Donem]]," ",""))))'
        ),
        "GecikmeGun": f'IF(tblVekalet[[#This Row],[BorcId]]="","",{yas_f})',
        "ZamTutari": (
            'IF(tblVekalet[[#This Row],[BorcId]]="","",'
            'IF(tblVekalet[[#This Row],[GecikmeGun]]>hbd_sureEsikGun,'
            'ROUND(tblVekalet[[#This Row],[BorcTutari]]*hbd_gecikmeZammi,2),0))'
        ),
        "ToplamBorc": (
            'IF(tblVekalet[[#This Row],[BorcId]]="","",'
            'tblVekalet[[#This Row],[BorcTutari]]+tblVekalet[[#This Row],[ZamTutari]])'
        ),
        "MotorAdim": (
            'IF(tblVekalet[[#This Row],[BorcId]]="","",'
            '(tblVekalet[[#This Row],[ZamTutari]]>0)+1)'
        ),
    }
    for i, (bid, mid, don, tut, vade) in enumerate(VEKALET_SATIR):
        r = ILK + i
        _sari(ws, r, 1, bid, baslik="Borç ID", mesaj="Benzersiz borç kimliği")
        _sari(ws, r, 2, mid, baslik="Müvekkil ID", mesaj="PAYDASLAR'daki kimlik")
        _sari(ws, r, 3, don, baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 4, tut, sayi=TL, baslik="Borç tutarı", mesaj="Vekalet ücreti TL")
        _sari(ws, r, 5, vade, sayi=TARİH, baslik="Vade", mesaj="GG.AA.YYYY")
    tablo_ekle(ws, "tblVekalet", f"A{HDR}:J{SON}", sut, formuller=form)
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Borç", mesaj="≥ 0 TL",
              hata_baslik="Borç", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "date", "2020-01-01", f"E{ILK}:E{SON}",
              baslik="Vade", mesaj="Geçerli tarih",
              hata_baslik="Tarih", hata_mesaj="Tarih girin",
              isaret="greaterThanOrEqual", f2="2035-12-31")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 11)})


def tahsilat(ws):
    sayfa_hazirla(ws, "TAHSILAT", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Tahsilat girişi — vekalet eşleştirme motoruna besleme", kalin=True, yazi=KOYU_LACIVERT)
    sut = ["OdemeId", "MuvekkilId", "Donem", "OdemeTutari", "OdemeTarihi",
           "HamAnahtar", "KaynakKartId", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "HamAnahtar": (
            'IF(tblTahsilat[[#This Row],[MuvekkilId]]="","",'
            'UPPER(TRIM(SUBSTITUTE(tblTahsilat[[#This Row],[MuvekkilId]]&'
            'tblTahsilat[[#This Row],[Donem]]," ",""))))'
        ),
        "KaynakKartId": 'IF(tblTahsilat[[#This Row],[OdemeId]]="","",tblTahsilat[[#This Row],[MuvekkilId]])',
        "MotorAdim": 'IF(tblTahsilat[[#This Row],[OdemeId]]="","",1)',
    }
    for i, (oid, mid, don, tut, tar) in enumerate(TAHSIL_SATIR):
        r = ILK + i
        _sari(ws, r, 1, oid, baslik="Ödeme ID", mesaj="Benzersiz ödeme")
        _sari(ws, r, 2, mid, baslik="Müvekkil ID", mesaj="Müvekkil kimliği")
        _sari(ws, r, 3, don, baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 4, tut, sayi=TL, baslik="Ödeme tutarı", mesaj="TL")
        _sari(ws, r, 5, tar, sayi=TARİH, baslik="Ödeme tarihi", mesaj="GG.AA.YYYY")
    tablo_ekle(ws, "tblTahsilat", f"A{HDR}:H{SON}", sut, formuller=form)
    dogrulama(ws, "decimal", "0", f"D{ILK}:D{SON}",
              baslik="Ödeme", mesaj="≥ 0",
              hata_baslik="Ödeme", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "date", "2020-01-01", f"E{ILK}:E{SON}",
              baslik="Tarih", mesaj="Geçerli tarih",
              hata_baslik="Tarih", hata_mesaj="Tarih girin",
              isaret="greaterThanOrEqual", f2="2035-12-31")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    h(ws, 3, 1, "Normalize müvekkil+dönem anahtarı + duplike (E01/E05)", kalin=True,
      yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "A Tarafı (Vekalet)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",UPPER(TRIM(SUBSTITUTE(tblAnahtarA[[#This Row],[Ham]]," ",""))))',
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": (
            'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
            'LEFT(tblAnahtarA[[#This Row],[Norm]],hbd_ikincilUzunluk))'
        ),
        "KaynakTaraf": '"A"',
        "Duplike": (
            'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
            'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))'
        ),
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(min(len(VEKALET_SATIR) + 5, 80)):
        r = ILK + i
        h(ws, r, 1, f'=IF(VEKALET!A{r}="","",VEKALET!F{r})')
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Tahsilat)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ",""))))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": (
            'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
            'LEFT(tblAnahtarB[[#This Row],[NormB]],hbd_ikincilUzunluk))'
        ),
        "KaynakB": '"B"',
        "DuplikeB": (
            'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
            'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))'
        ),
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(min(len(TAHSIL_SATIR) + 5, 80)):
        r = ILK + i
        h(ws, r, 9, f'=IF(TAHSILAT!A{r}="","",TAHSILAT!F{r})')
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 16)})
    ws.column_dimensions["I"].width = 22


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    h(ws, 3, 1, "Dört kova — vekalet↔tahsilat (E02/E03)", kalin=True, yazi=KOYU_LACIVERT)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "NetKatki", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",COUNTIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]]))',
        "B_Tutar": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'SUMIF(tblAnahtarB[NormB],tblEslesme[[#This Row],[A_Norm]],tblTahsilat[OdemeTutari]))'
        ),
        "A_Tutar": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'SUMIF(tblAnahtarA[Norm],tblEslesme[[#This Row],[A_Norm]],tblVekalet[ToplamBorc]))'
        ),
        "Fark": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))'
        ),
        "AbsFark": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",ABS(tblEslesme[[#This Row],[Fark]]))'
        ),
        "Kova": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]=0,"eslesmedi",'
            'IF(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)=ROUND(tblEslesme[[#This Row],[B_Tutar]],2),"eslesti",'
            'IF(ABS(ROUND(tblEslesme[[#This Row],[A_Tutar]],2)-ROUND(tblEslesme[[#This Row],[B_Tutar]],2))<=toleransKurus,'
            '"tolerans_icinde","tutar_farki"))))'
        ),
        "IkincilEslesme": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(tblEslesme[[#This Row],[B_Sayim]]>0,"BIRINCIL",'
            'IF(COUNTIF(tblAnahtarB[IkincilB],tblAnahtarA[[#This Row],[IkincilAnahtar]])>0,'
            '"IKINCIL","YOK")))'
        ),
        "DuplikeBayrak": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblAnahtarA[[#This Row],[Duplike]])'
        ),
        "FarkOran": (
            'IF(OR(tblEslesme[[#This Row],[A_Norm]]="",tblEslesme[[#This Row],[A_Tutar]]=0),0,'
            'tblEslesme[[#This Row],[AbsFark]]/ABS(tblEslesme[[#This Row],[A_Tutar]]))'
        ),
        "YuvarlakA": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[A_Tutar]],2))'
        ),
        "YuvarlakB": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",ROUND(tblEslesme[[#This Row],[B_Tutar]],2))'
        ),
        "ToleransTest": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            '--(tblEslesme[[#This Row],[AbsFark]]<=toleransKurus))'
        ),
        "EsikUstu": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            '--(tblEslesme[[#This Row],[AbsFark]]>hbd_esikFark))'
        ),
        "NetKatki": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'IF(OR(tblEslesme[[#This Row],[Kova]]="eslesti",'
            'tblEslesme[[#This Row],[Kova]]="tolerans_icinde"),'
            'tblEslesme[[#This Row],[A_Tutar]],0))'
        ),
        "MotorAdim": (
            'IF(tblEslesme[[#This Row],[A_Norm]]="","",'
            'tblEslesme[[#This Row],[B_Sayim]]+tblEslesme[[#This Row],[ToleransTest]]'
            '+tblEslesme[[#This Row],[EsikUstu]])'
        ),
    }
    tablo_ekle(ws, "tblEslesme", f"A{HDR}:P{SON}", sutunlar, formuller=formuller)
    h(ws, 2, 1, "Kova bütünlük", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 2, 3, dort_kova_butunluk(
        'COUNTA(tblEslesme[A_Norm])',
        'COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde")'
        '+COUNTIF(tblEslesme[Kova],"tutar_farki")+COUNTIF(tblEslesme[Kova],"eslesmedi")',
    ))
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 17)})


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    from openpyxl.styles import Alignment
    h(ws, 3, 1, "Kök neden — eşleşmeyen fark; TANIMSIZ elle incele (E04)", kalin=True,
      yazi=KOYU_LACIVERT)
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.column_dimensions["A"].width = 58
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "EKSIK_ODEME", "test": "fark>0",
         "aciklama": "Vekalet tahsilattan büyük — eksik ödeme"},
        {"neden_kodu": "FAZLA_ODEME", "test": "fark<0",
         "aciklama": "Tahsilat vekaletten büyük — fazla/yanlış dönem"},
        {"neden_kodu": "DONEM_UYUSMAZ", "test": "ikincil=YOK",
         "aciklama": "Dönem anahtarı eşleşmedi"},
        {"neden_kodu": "DUPLIKE", "test": "duplike=EVET",
         "aciklama": "Çift kayıt şüphesi"},
    ])
    h(ws, 4, 8, "Sıra", kalin=True)
    h(ws, 4, 9, "Neden kodu", kalin=True)
    h(ws, 4, 10, "Test", kalin=True)
    h(ws, 4, 11, "Açıklama", kalin=True)
    for i, k in enumerate(kurallar):
        r = 5 + i
        h(ws, r, 8, k["sira"], sayi="0")
        h(ws, r, 9, k["neden_kodu"])
        h(ws, r, 10, k["test_formul_sablon"])
        h(ws, r, 11, k["aciklama"], kaydir=True)

    sut = ["Norm", "Fark", "Kova", "NedenKodu", "Aciklama", "ElleIncele"]
    baslik_satiri(ws, HDR, sut)
    form_yazi = {
        "Norm": (
            'IFERROR(IF(INDEX(tblEslesme[A_Norm],ROW()-ROW($A$5))="","",'
            'INDEX(tblEslesme[A_Norm],ROW()-ROW($A$5))),"")'
        ),
        "Fark": (
            'IF(tblKok[[#This Row],[Norm]]="","",'
            'IFERROR(INDEX(tblEslesme[Fark],ROW()-ROW($A$5)),0))'
        ),
        "Kova": (
            'IF(tblKok[[#This Row],[Norm]]="","",'
            'IFERROR(INDEX(tblEslesme[Kova],ROW()-ROW($A$5)),""))'
        ),
        "NedenKodu": (
            'IF(tblKok[[#This Row],[Norm]]="","",'
            'IF(tblKok[[#This Row],[Kova]]="eslesti","",'
            'IF(tblKok[[#This Row],[Kova]]="tolerans_icinde","",'
            'IF(tblKok[[#This Row],[Fark]]>0,"EKSIK_ODEME",'
            'IF(tblKok[[#This Row],[Fark]]<0,"FAZLA_ODEME","TANIMSIZ")))))'
        ),
        "Aciklama": (
            'IF(tblKok[[#This Row],[NedenKodu]]="","",'
            'IF(tblKok[[#This Row],[NedenKodu]]="TANIMSIZ","Elle incele — TANIMSIZ kova",'
            'IF(tblKok[[#This Row],[NedenKodu]]="EKSIK_ODEME","Eksik ödeme / gecikme",'
            'IF(tblKok[[#This Row],[NedenKodu]]="FAZLA_ODEME","Fazla veya yanlış dönem","İncele"))))'
        ),
        "ElleIncele": (
            'IF(tblKok[[#This Row],[NedenKodu]]="TANIMSIZ","EVET","HAYIR")'
        ),
    }
    tablo_ekle(ws, "tblKok", f"A{HDR}:F{SON}", sut, formuller=form_yazi)
    genislik(ws, {"B": 12, "C": 16, "D": 16, "E": 36, "F": 12,
                  "H": 8, "I": 14, "J": 14, "K": 36})
    sabitle(ws, f"A{ILK}")


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Dosya durum makinesi — yetim / geçersiz geçiş / ZİNCİR KIRIK / süre", kalin=True,
      yazi=KOYU_LACIVERT)
    sut = [
        "IslemId", "KaynakKartId", "BorcId", "VadeTarihi", "OdemeTarihi",
        "OncekiDurum", "Durum", "BorcTutari", "OdemeTutari", "Tutar",
        "GecikmeGun", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sut)
    yas_f = yaslandirma_gun("tblAkis[[#This Row],[VadeTarihi]]", "raporTarihi").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblPaydas[MuvekkilId]").lstrip("=")
    gecis_birlesik = 'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[BorcTutari]]",
        "tblAkis[[#This Row],[OdemeTutari]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_guvenli = (
        f'IF(OR(tblAkis[[#This Row],[Durum]]="KAPALI",'
        f'tblAkis[[#This Row],[OdemeTutari]]>0),{zincir_f},"OK")'
    )
    form = {
        "GecikmeGun": (
            f'IF(tblAkis[[#This Row],[IslemId]]="","",'
            f'IF(OR(tblAkis[[#This Row],[Durum]]="SONUC",'
            f'tblAkis[[#This Row],[Durum]]="KAPALI"),0,{yas_f}))'
        ),
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_guvenli})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
        "Tutar": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[BorcTutari]]-tblAkis[[#This Row],[OdemeTutari]])'
        ),
    }

    # demo akış satırları
    durumlar = ["ACIK", "VEKALET", "DAVADA", "DURUSMA", "SONUC", "KAPALI", "VEKALET", "DAVADA"]
    oncekiler = ["", "ACIK", "VEKALET", "DAVADA", "DURUSMA", "SONUC", "ACIK", "VEKALET"]
    for i in range(DEMO_AKIS):
        r = ILK + i
        mid = MUVEKKILLER[i % len(MUVEKKILLER)][0]
        if i == 15:
            mid = "MUV-999"  # yetim
        d = durumlar[i % len(durumlar)]
        o = oncekiler[i % len(oncekiler)]
        if i == 10:
            o, d = "ACIK", "DAVADA"  # geçersiz geçiş
        borc = 350.0 + i * 20
        odeme = borc if d in ("KAPALI", "SONUC") else (0 if i != 12 else borc + 400)
        bid = VEKALET_SATIR[i % len(VEKALET_SATIR)][0]
        vade = date(2026, 7, 10) + timedelta(days=i)
        _sari(ws, r, 1, f"DOS-{i+1:03d}", baslik="İşlem ID", mesaj="Dosya işlem kimliği")
        _sari(ws, r, 2, mid, baslik="Kaynak kart", mesaj="Müvekkil ID (boş=yetim)")
        _sari(ws, r, 3, bid, baslik="Borç ID", mesaj="VEKALET satır referansı")
        _sari(ws, r, 4, vade, sayi=TARİH, baslik="Vade", mesaj="Vade tarihi")
        _sari(ws, r, 5, date(2026, 8, 1) if odeme else None, sayi=TARİH,
              baslik="Ödeme tarihi", mesaj="Varsa ödeme günü")
        _sari(ws, r, 6, o or None, baslik="Önceki durum", mesaj="durum_id")
        _sari(ws, r, 7, d, baslik="Durum", mesaj="Durum haritasından")
        _sari(ws, r, 8, borc, sayi=TL, baslik="Borç tutarı", mesaj="TL")
        _sari(ws, r, 9, odeme, sayi=TL, baslik="Ödeme tutarı", mesaj="TL")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:O{SON}", sut, formuller=form)
    dogrulama(ws, "list", "ListeDurum", f"G{ILK}:G{SON}",
              baslik="Durum", mesaj="Durum seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "decimal", "0", f"H{ILK}:I{SON}",
              baslik="Tutar", mesaj="≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 16)})


def gider(ws):
    sayfa_hazirla(ws, "GIDER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Gider takibi — bütçe vs gerçekleşen", kalin=True, yazi=KOYU_LACIVERT)
    sut = ["GiderId", "Kalem", "Butce", "Gerceklesen", "Donem", "Sapma", "SapmaOran", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "Sapma": (
            'IF(tblGider[[#This Row],[GiderId]]="","",'
            'tblGider[[#This Row],[Gerceklesen]]-tblGider[[#This Row],[Butce]])'
        ),
        "SapmaOran": (
            'IF(OR(tblGider[[#This Row],[GiderId]]="",tblGider[[#This Row],[Butce]]=0),0,'
            'tblGider[[#This Row],[Sapma]]/ABS(tblGider[[#This Row],[Butce]]))'
        ),
        "MotorAdim": 'IF(tblGider[[#This Row],[GiderId]]="","",1)',
    }
    for i, (gid, kalem, butce, ger, don) in enumerate(GIDER_SATIR):
        r = ILK + i
        _sari(ws, r, 1, gid, baslik="Gider ID", mesaj="Benzersiz")
        _sari(ws, r, 2, kalem, baslik="Kalem", mesaj="Gider kalemi")
        _sari(ws, r, 3, butce, sayi=TL, baslik="Bütçe", mesaj="Plan TL")
        _sari(ws, r, 4, ger, sayi=TL, baslik="Gerçekleşen", mesaj="Fiili TL")
        _sari(ws, r, 5, don, baslik="Dönem", mesaj="YYYY-AA")
    tablo_ekle(ws, "tblGider", f"A{HDR}:H{SON}", sut, formuller=form)
    dogrulama(ws, "decimal", "0", f"C{ILK}:D{SON}",
              baslik="Tutar", mesaj="≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["B"].width = 18


def kararlar(ws):
    sayfa_hazirla(ws, "KARARLAR", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Genel kurul / yönetim kararları — kanıt niteliğinde defter", kalin=True,
      yazi=KOYU_LACIVERT)
    sut = ["KararId", "Tarih", "Konu", "Sonuc", "Organ", "MotorAdim"]
    baslik_satiri(ws, HDR, sut)
    form = {"MotorAdim": 'IF(tblKararlar[[#This Row],[KararId]]="","",1)'}
    for i, (kid, tar, konu, son, organ) in enumerate(KARAR_SATIR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Karar ID", mesaj="Benzersiz")
        _sari(ws, r, 2, tar, sayi=TARİH, baslik="Tarih", mesaj="Karar günü")
        _sari(ws, r, 3, konu, baslik="Konu", mesaj="Gündem maddesi")
        _sari(ws, r, 4, son, baslik="Sonuç", mesaj="ONAY/ERTELE/TASLAK")
        _sari(ws, r, 5, organ, baslik="Organ", mesaj="BURO veya MUVEKKIL")
    tablo_ekle(ws, "tblKararlar", f"A{HDR}:F{SON}", sut, formuller=form)
    dogrulama(ws, "list", "ListeKararSonuc", f"D{ILK}:D{SON}",
              baslik="Sonuç", mesaj="Listeden",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeOrgan", f"E{ILK}:E{SON}",
              baslik="Organ", mesaj="BURO veya MUVEKKIL",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {"A": 12, "B": 12, "C": 28, "D": 12, "E": 10, "F": 12})


def pano(ws):
    from openpyxl.styles import Alignment
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=16)
    h(ws, 3, 1, "Büro panosu — kasa, vekalet tahsilatı, dosya kuyruğu, süre", kalin=True, boyut=14,
      yazi=KOYU_LACIVERT)
    ws["A3"].alignment = Alignment(wrap_text=True)
    h(ws, 4, 1, "Karar kapısı")
    h(ws, 5, 1, "Karar")
    h(ws, 5, 2, "=hbd_kararKapi", kalin=True, boyut=16)

    kpis = [
        (7, "Toplam borç", "=hbd_toplamBorc", TL),
        (8, "Toplam tahsilat", "=hbd_toplamTahsilat", TL),
        (9, "Tahsilat oranı", "=hbd_tahsilatOrani", YÜZDE),
        (10, "Kasa bakiyesi", "=hbd_kasaBakiye", TL),
        (11, "Geciken müvekkil", "=hbd_gecikenUye", "0"),
        (12, "Süre/dosya kuyruk", "=hbd_sureKuyruk", "0"),
        (13, "Yetim kayıt", "=hbd_yetimKayit", "0"),
        (14, "Geçersiz geçiş", "=hbd_gecersizGecis", "0"),
        (15, "ZİNCİR KIRIK", "=hbd_zincirKirik", "0"),
        (16, "Eşleşme oranı", "=hbd_eslesmeOrani", YÜZDE),
        (17, "Gider sapması", "=hbd_giderSapma", TL),
        (18, "Veri kalite", "=hbd_modulO8", "0"),
        (19, "Kova bütünlük", "=hbd_kovaButunluk", None),
        (20, "Karar sayısı", "=hbd_kararSayisi", "0"),
        (21, "Eşleşen kova", '=COUNTIF(tblEslesme[Kova],"eslesti")', "0"),
        (22, "Duruşma SUMIF", '=SUMIF(tblAkis[Durum],"DURUSMA",tblAkis[Tutar])', TL),
    ]
    for r, ad, formul, sayi in kpis:
        h(ws, r, 1, ad)
        h(ws, r, 2, formul, sayi=sayi, kalin=True)

    # trend — B formül (canlı), C sabit demo (grafik D09b)
    h(ws, 7, 4, "Aylık vekalet tahsilat trendi", kalin=True, yazi=KOYU_LACIVERT)
    aylar = [
        ("Ocak", "hbd_trendOcak", 0.78), ("Şubat", "hbd_trendSubat", 0.80),
        ("Mart", "hbd_trendMart", 0.82), ("Nisan", "hbd_trendNisan", 0.85),
        ("Mayıs", "hbd_trendMayis", 0.88), ("Haziran", "hbd_trendHaziran", 0.90),
        ("Temmuz", "hbd_trendTemmuz", 0.95), ("Ağustos", "hbd_trendAgustos", 1.00),
        ("Eylül", "hbd_trendEylul", 0.97), ("Ekim", "hbd_trendEkim", 0.93),
        ("Kasım", "hbd_trendKasim", 0.90), ("Aralık", "hbd_trendAralik", 0.88),
    ]
    h(ws, 8, 4, "Ay", kalin=True)
    h(ws, 8, 5, "CanliEndeks", kalin=True)
    h(ws, 8, 6, "DemoEndeks", kalin=True)
    h(ws, 8, 7, "DemoTutar", kalin=True)
    demo_baz = 28000.0
    for i, (ay, ad, demo) in enumerate(aylar):
        r = 9 + i
        h(ws, r, 4, ay)
        h(ws, r, 5, f"=IFERROR(hbd_toplamTahsilat*{ad}/MAX(hbd_trendAgustos,1),0)", sayi=TL)
        h(ws, r, 6, demo, sayi="0.00")
        h(ws, r, 7, round(demo_baz * demo, 2), sayi=TL)

    h(ws, 7, 9, "Senaryo", kalin=True)
    h(ws, 8, 9, "Baz tahsilat")
    h(ws, 8, 10, "=hbd_toplamTahsilat", sayi=TL)
    h(ws, 9, 9, "İyimser")
    h(ws, 9, 10, "=hbd_toplamTahsilat*hbd_senaryoIyimser", sayi=TL)
    h(ws, 10, 9, "Kötümser")
    h(ws, 10, 10, "=hbd_toplamTahsilat*hbd_senaryoKotumser", sayi=TL)
    h(ws, 12, 9, "Tornado 1 (zam)")
    h(ws, 12, 10, "=hbd_modulO2", sayi=TL)
    h(ws, 13, 9, "P90 gecikme")
    h(ws, 13, 10, "=hbd_modulI2", sayi="0.0")

    c1 = BarChart()
    c1.title = "Aylık endeks"
    c1.style = 10
    data = Reference(ws, min_col=6, min_row=8, max_row=20)
    cats = Reference(ws, min_col=4, min_row=9, max_row=20)
    c1.add_data(data, titles_from_data=True)
    c1.set_categories(cats)
    c1.width = 12
    c1.height = 8
    ws.add_chart(c1, "A22")

    c2 = LineChart()
    c2.title = "Trend tutar"
    data2 = Reference(ws, min_col=7, min_row=8, max_row=20)
    c2.add_data(data2, titles_from_data=True)
    c2.set_categories(cats)
    c2.width = 12
    c2.height = 8
    ws.add_chart(c2, "H22")

    c3 = PieChart()
    c3.title = "Kuyruk dağılımı"
    h(ws, 22, 4, "Etiket", kalin=True)
    h(ws, 22, 5, "Adet", kalin=True)
    h(ws, 23, 4, "Kuyruk ACIK")
    h(ws, 23, 5, 4, sayi="0")
    h(ws, 24, 4, "Kuyruk VEKALET")
    h(ws, 24, 5, 5, sayi="0")
    h(ws, 25, 4, "Kuyruk DURUSMA")
    h(ws, 25, 5, 5, sayi="0")
    h(ws, 26, 4, "Kuyruk KAPALI")
    h(ws, 26, 5, 4, sayi="0")
    # Canlı kuyruk (A05) — grafik serisinden ayrı, D09b sıfır-seri tuzağı yok
    h(ws, 27, 4, "Canli ACIK")
    h(ws, 27, 5, '=COUNTIF(tblAkis[Durum],"ACIK")', sayi="0")
    h(ws, 28, 4, "Canli VEKALET")
    h(ws, 28, 5, '=COUNTIF(tblAkis[Durum],"VEKALET")', sayi="0")
    h(ws, 29, 4, "Canli DURUSMA")
    h(ws, 29, 5, '=COUNTIF(tblAkis[Durum],"DURUSMA")', sayi="0")
    h(ws, 30, 4, "Canli KAPALI")
    h(ws, 30, 5, '=COUNTIF(tblAkis[Durum],"KAPALI")', sayi="0")
    pdata = Reference(ws, min_col=5, min_row=22, max_row=26)
    pcats = Reference(ws, min_col=4, min_row=23, max_row=26)
    c3.add_data(pdata, titles_from_data=True)
    c3.set_categories(pcats)
    c3.width = 10
    c3.height = 8
    ws.add_chart(c3, "A38")

    c4 = BarChart()
    c4.title = "Gider sapması"
    h(ws, 22, 9, "Kalem", kalin=True)
    h(ws, 22, 10, "Sapma", kalin=True)
    for i, g in enumerate(GIDER_SATIR[:6]):
        h(ws, 23 + i, 9, f"Sapma {g[1]}")
        h(ws, 23 + i, 10, g[3] - g[2], sayi=TL)
    gdata = Reference(ws, min_col=10, min_row=22, max_row=28)
    gcats = Reference(ws, min_col=9, min_row=23, max_row=28)
    c4.add_data(gdata, titles_from_data=True)
    c4.set_categories(gcats)
    c4.width = 12
    c4.height = 8
    ws.add_chart(c4, "H38")

    c5 = BarChart()
    c5.title = "Senaryo bandı"
    h(ws, 30, 9, "Senaryo", kalin=True)
    h(ws, 30, 10, "Tutar", kalin=True)
    h(ws, 31, 9, "Kötümser")
    h(ws, 31, 10, 23800.0, sayi=TL)
    h(ws, 32, 9, "Baz")
    h(ws, 32, 10, 28000.0, sayi=TL)
    h(ws, 33, 9, "İyimser")
    h(ws, 33, 10, 30240.0, sayi=TL)
    sdata = Reference(ws, min_col=10, min_row=30, max_row=33)
    scats = Reference(ws, min_col=9, min_row=31, max_row=33)
    c5.add_data(sdata, titles_from_data=True)
    c5.set_categories(scats)
    c5.width = 10
    c5.height = 7
    ws.add_chart(c5, "A52")

    c6 = LineChart()
    c6.title = "Kalite / oran"
    h(ws, 38, 4, "Metrik", kalin=True)
    h(ws, 38, 5, "Deger", kalin=True)
    h(ws, 39, 4, "Tahsilat")
    h(ws, 39, 5, 0.88, sayi=YÜZDE)
    h(ws, 40, 4, "Eşleşme")
    h(ws, 40, 5, 0.75, sayi=YÜZDE)
    h(ws, 41, 4, "Kalite")
    h(ws, 41, 5, 0.82, sayi=YÜZDE)
    mdata = Reference(ws, min_col=5, min_row=38, max_row=41)
    mcats = Reference(ws, min_col=4, min_row=39, max_row=41)
    c6.add_data(mdata, titles_from_data=True)
    c6.set_categories(mcats)
    c6.width = 10
    c6.height = 7
    ws.add_chart(c6, "H52")

    c7 = BarChart()
    c7.title = "Uyarı sayıları"
    h(ws, 44, 4, "Uyarı", kalin=True)
    h(ws, 44, 5, "Adet", kalin=True)
    h(ws, 45, 4, "Yetim")
    h(ws, 45, 5, 1, sayi="0")
    h(ws, 46, 4, "Geçersiz")
    h(ws, 46, 5, 1, sayi="0")
    h(ws, 47, 4, "Zincir")
    h(ws, 47, 5, 1, sayi="0")
    udata = Reference(ws, min_col=5, min_row=44, max_row=47)
    ucats = Reference(ws, min_col=4, min_row=45, max_row=47)
    c7.add_data(udata, titles_from_data=True)
    c7.set_categories(ucats)
    c7.width = 10
    c7.height = 7
    ws.add_chart(c7, "A64")

    c8 = PieChart()
    c8.title = "Karar tarafı"
    h(ws, 44, 9, "Organ", kalin=True)
    h(ws, 44, 10, "Adet", kalin=True)
    h(ws, 45, 9, "BURO")
    h(ws, 45, 10, 4, sayi="0")
    h(ws, 46, 9, "MUVEKKIL")
    h(ws, 46, 10, 2, sayi="0")
    odata = Reference(ws, min_col=10, min_row=44, max_row=46)
    ocats = Reference(ws, min_col=9, min_row=45, max_row=46)
    c8.add_data(odata, titles_from_data=True)
    c8.set_categories(ocats)
    c8.width = 10
    c8.height = 7
    ws.add_chart(c8, "H64")

    baski_hazirla(ws, "A1:N70", f"{URUN_AD} | PANO | {SURUM}")
    genislik(ws, {"A": 20, "B": 16, "D": 24, "E": 14, "F": 12, "G": 14,
                  "I": 26, "J": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Karar kapısı ve gerekçe", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Karar")
    h(ws, 5, 2, "=hbd_kararKapi", kalin=True, boyut=18)
    h(ws, 6, 1, "Gerekçe")
    h(ws, 6, 2, "=hbd_kararGerekce", kaydir=True)
    ws.merge_cells("B6:H6")
    h(ws, 8, 1, "Tahsilat oranı")
    h(ws, 8, 2, "=hbd_tahsilatOrani", sayi=YÜZDE)
    h(ws, 9, 1, "Hedef")
    h(ws, 9, 2, "=GIRDI!B9", sayi=YÜZDE)
    h(ws, 10, 1, "Zincir kırık")
    h(ws, 10, 2, "=hbd_zincirKirik", sayi="0")
    h(ws, 11, 1, "Yetim")
    h(ws, 11, 2, "=hbd_yetimKayit", sayi="0")
    h(ws, 13, 1, "Modül özeti", kalin=True, yazi=KOYU_LACIVERT)
    moduller = [
        (14, "T1 Tahsilat oranı", "=hbd_modulT1", YÜZDE, "=hbd_modulT1Yorum"),
        (15, "T2 Trend Ağustos", "=hbd_modulT2", "0.00", "=hbd_modulT2Yorum"),
        (16, "O1 Geciken", "=hbd_modulO1", "0", "=hbd_modulO1Yorum"),
        (17, "O2 Tornado", "=hbd_modulO2", TL, "=hbd_modulO2Yorum"),
        (18, "O3 Senaryo bandı", "=hbd_modulO3", TL, "=hbd_modulO3Yorum"),
        (19, "O6 Gider HHI", "=hbd_modulO6", "0.00", "=hbd_modulO6Yorum"),
        (20, "O8 Kalite", "=hbd_modulO8", "0", "=hbd_modulO8Yorum"),
        (21, "I1 Tahmin", "=hbd_modulI1", TL, "=hbd_modulI1Yorum"),
        (22, "I2 P90", "=hbd_modulI2", "0.0", "=hbd_modulI2Yorum"),
    ]
    for r, ad, val, sayi, yorum in moduller:
        h(ws, r, 1, ad)
        h(ws, r, 2, val, sayi=sayi)
        h(ws, r, 3, yorum, kaydir=True)
    genislik(ws, {"A": 22, "B": 18, "C": 48})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    baski_hazirla(ws, "A1:H28", f"{URUN_AD} — Genel Kurul Kanıt Raporu")
    h(ws, 2, 1, "MÜVEKKİL / BÜRO DÖNEM KANIT RAPORU", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 1, f"Ürün: {URUN_AD} · Sürüm {SURUM}")
    h(ws, 4, 1, "Rapor tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Kurum")
    h(ws, 5, 2, "=GIRDI!B5")
    h(ws, 6, 1, "Karar")
    h(ws, 6, 2, "=hbd_kararKapi", kalin=True)
    h(ws, 8, 1, "Özet KPI", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, f, s) in enumerate([
        ("Toplam borç", "=hbd_toplamBorc", TL),
        ("Toplam tahsilat", "=hbd_toplamTahsilat", TL),
        ("Tahsilat oranı", "=hbd_tahsilatOrani", YÜZDE),
        ("Kasa bakiyesi", "=hbd_kasaBakiye", TL),
        ("Gider sapması", "=hbd_giderSapma", TL),
        ("Süre/dosya kuyruk", "=hbd_sureKuyruk", "0"),
    ], 9):
        h(ws, i, 1, ad)
        h(ws, i, 2, f, sayi=s)
    h(ws, 16, 1, "Varsayımlar", kalin=True)
    h(ws, 17, 1, "Gecikme zammı / süre eşiği AYARLAR.hbd_gecikmeZammi (AAÜT/HMK çerçevesi). "
      "Bu çıktı hukuki tavsiye değildir; karar destek kanıtıdır.", kaydir=True)
    ws.merge_cells("A17:H18")
    h(ws, 20, 1, "İmza — Avukat", kalin=True)
    h(ws, 20, 3, "İmza — Müvekkil", kalin=True)
    h(ws, 22, 1, "________________")
    h(ws, 22, 3, "________________")
    h(ws, 24, 1, "Uyarılar: Yetim / geçersiz geçiş / ZİNCİR KIRIK sayıları PANO'da.", yazi=GRİ)
    genislik(ws, {"A": 48, "B": 18, "C": 24})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Çalışma kitabı sağlık paneli", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    testler = [
        (5, "Müvekkil dolu mu?", '=IF(COUNTA(tblPaydas[MuvekkilId])=0,"BOŞ","DOLU")'),
        (6, "Dosya dolu mu?", '=IF(COUNTA(tblDosya[DosyaId])=0,"BOŞ","DOLU")'),
        (7, "Vekalet satırı var mı?", '=IF(COUNTA(tblVekalet[BorcId])=0,"BOŞ","DOLU")'),
        (8, "Tahsilat var mı?", '=IF(COUNTA(tblTahsilat[OdemeId])=0,"BOŞ","DOLU")'),
        (9, "Kova bütünlük", "=hbd_kovaButunluk"),
        (10, "Yetim var mı?", '=IF(hbd_yetimKayit>0,"HATALI","TEMİZ")'),
        (11, "Geçersiz geçiş", '=IF(hbd_gecersizGecis>0,"HATALI","TEMİZ")'),
        (12, "Zincir kırık", '=IF(hbd_zincirKirik>0,"ALARM","TEMİZ")'),
        (13, "Karar kapısı", "=hbd_kararKapi"),
    ]
    for r, ad, f in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, f)
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Demo veri özeti — dosya/vekalet/tahsilat", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Müvekkil sayısı")
    h(ws, 5, 2, len(MUVEKKILLER), sayi="0")
    h(ws, 6, 1, "Dosya sayısı")
    h(ws, 6, 2, len(DOSYA_SATIR), sayi="0")
    h(ws, 7, 1, "Vekalet satırı")
    h(ws, 7, 2, len(VEKALET_SATIR), sayi="0")
    h(ws, 8, 1, "Tahsilat satırı")
    h(ws, 8, 2, len(TAHSIL_SATIR), sayi="0")
    h(ws, 9, 1, "Dosya akış satırı")
    h(ws, 9, 2, DEMO_AKIS, sayi="0")
    h(ws, 10, 1, "Not: Sarı hücrelere kendi kurum verinizi yazın; formüller otomatik güncellenir.",
      kaydir=True, yazi=GRİ)
    genislik(ws, {"A": 22, "B": 12})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Açılır listeler", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "EvetHayir", kalin=True)
    h(ws, 6, 1, "EVET")
    h(ws, 7, 1, "HAYIR")
    h(ws, 5, 3, "Durum", kalin=True)
    for i, d in enumerate(DURUM_HARITASI):
        h(ws, 6 + i, 3, d["durum_id"])
    h(ws, 5, 5, "KararSonuc", kalin=True)
    for i, s in enumerate(["ONAY", "ERTELE", "TASLAK", "RED"]):
        h(ws, 6 + i, 5, s)
    h(ws, 5, 7, "Organ", kalin=True)
    h(ws, 6, 7, "BURO")
    h(ws, 7, 7, "MUVEKKIL")
    h(ws, 5, 9, "IzinliGecis", kalin=True)
    izinler = []
    for d in DURUM_HARITASI:
        once = d["durum_id"]
        sonrakiler = [s for s in (d["izinli_sonraki_durumlar"] or "").split("|") if s]
        izinler.extend(f"{once}|{son}" for son in sonrakiler)
    for i, g in enumerate(izinler):
        h(ws, 6 + i, 9, g)
    genislik(ws, {"A": 12, "C": 14, "E": 12, "G": 10, "I": 18})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Tek kaynak parametreler + durum haritası (A01/A06)", kalin=True,
      yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "anahtar", kalin=True)
    h(ws, 4, 2, "deger", kalin=True)
    h(ws, 4, 3, "birim", kalin=True)
    h(ws, 4, 4, "kaynak", kalin=True)
    h(ws, 4, 5, "yururluk_tarihi", kalin=True)
    h(ws, 4, 6, "aciklama", kalin=True)

    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Yönetim kurulu", "01.01.2026", "Rapor tarihi", TARİH),
        ("toleransKurus", 0.05, "TL", "Mutabakat politikası", "01.01.2026", "Eşleştirme toleransı", "0.00"),
        ("zincirTolerans", 100.0, "TL", "Dosya süre prosedürü", "01.01.2026", "ZİNCİR KIRIK eşiği", TL),
        ("hbd_gecikmeZammi", 0.05, "oran", "AAÜT/HMK", "01.01.2026", "Gecikme zammı", YÜZDE),
        ("hbd_sureEsikGun", 30, "gun", "AAÜT/HMK", "01.01.2026", "Süre alarm gün eşiği", "0"),
        ("hbd_esikFark", 50.0, "TL", "Yönetim kurulu", "01.01.2026", "Fark alarm eşiği", TL),
        ("hbd_ikincilUzunluk", 12, "adet", "Anahtar stratejisi", "01.01.2026", "İkincil anahtar", "0"),
        ("hbd_hedefTahsilat", 0.92, "oran", "Bütçe kararı", "01.01.2026", "Hedef tahsilat", YÜZDE),
        ("hbd_senaryoIyimser", 1.08, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan", "0.00"),
        ("hbd_senaryoKotumser", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Kötümser çarpan", "0.00"),
        ("hbd_fiyatModeli", 9900, "TL", "Fiyat çapası D0", "01.01.2026", "Satış fiyatı", TL),
        ("hbd_olcekHedef", 50000, "satir", "SPEC olcek", "01.01.2026", "Ölçek sözleşmesi", "0"),
        ("hbd_trendOcak", 0.78, "oran", "Tarihsel tahsilat", "01.01.2026", "Ocak trend", "0.00"),
        ("hbd_trendSubat", 0.80, "oran", "Tarihsel tahsilat", "01.01.2026", "Subat trend", "0.00"),
        ("hbd_trendMart", 0.82, "oran", "Tarihsel tahsilat", "01.01.2026", "Mart trend", "0.00"),
        ("hbd_trendNisan", 0.85, "oran", "Tarihsel tahsilat", "01.01.2026", "Nisan trend", "0.00"),
        ("hbd_trendMayis", 0.88, "oran", "Tarihsel tahsilat", "01.01.2026", "Mayis trend", "0.00"),
        ("hbd_trendHaziran", 0.90, "oran", "Tarihsel tahsilat", "01.01.2026", "Haziran trend", "0.00"),
        ("hbd_trendTemmuz", 0.95, "oran", "Tarihsel tahsilat", "01.01.2026", "Temmuz trend", "0.00"),
        ("hbd_trendAgustos", 1.00, "oran", "Tarihsel tahsilat", "01.01.2026", "Agustos trend", "0.00"),
        ("hbd_trendEylul", 0.97, "oran", "Tarihsel tahsilat", "01.01.2026", "Eylul trend", "0.00"),
        ("hbd_trendEkim", 0.93, "oran", "Tarihsel tahsilat", "01.01.2026", "Ekim trend", "0.00"),
        ("hbd_trendKasim", 0.90, "oran", "Tarihsel tahsilat", "01.01.2026", "Kasim trend", "0.00"),
        ("hbd_trendAralik", 0.88, "oran", "Tarihsel tahsilat", "01.01.2026", "Aralik trend", "0.00"),
    ]
    for i, (ad, deger, birim, kaynak, yur, acik, sayi) in enumerate(params):
        r = 5 + i
        h(ws, r, 1, ad)
        _sari(ws, r, 2, deger, sayi=sayi, baslik=ad, mesaj=acik)
        h(ws, r, 3, birim)
        h(ws, r, 4, kaynak)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik)

    # durum haritası
    h(ws, 5, 7, "durum_id", kalin=True)
    h(ws, 5, 8, "durum_adi", kalin=True)
    h(ws, 5, 9, "sira", kalin=True)
    h(ws, 5, 10, "izinli_sonraki_durumlar", kalin=True)
    h(ws, 5, 11, "kategori", kalin=True)
    for i, d in enumerate(DURUM_HARITASI):
        r = 6 + i
        h(ws, r, 7, d["durum_id"])
        h(ws, r, 8, d["durum_adi"])
        h(ws, r, 9, d["sira"], sayi="0")
        h(ws, r, 10, d["izinli_sonraki_durumlar"])
        h(ws, r, 11, d["kategori"])  # acik/kapali — A06

    # ayna KPI / karar (formül)
    h(ws, 6, 13, "Ayna / karar motoru", kalin=True, yazi=KOYU_LACIVERT)
    aynalar = [
        (7, "hbd_toplamBorc", '=SUMIF(tblVekalet[BorcId],"<>",tblVekalet[ToplamBorc])', TL),
        (8, "hbd_toplamTahsilat", '=SUMIF(tblTahsilat[OdemeId],"<>",tblTahsilat[OdemeTutari])', TL),
        (9, "hbd_tahsilatOrani",
         '=IF(hbd_toplamBorc=0,0,hbd_toplamTahsilat/hbd_toplamBorc)', YÜZDE),
        (10, "hbd_kasaBakiye",
         "=hbd_toplamTahsilat-SUMIF(tblGider[GiderId],\"<>\",tblGider[Gerceklesen])", TL),
        (11, "hbd_gecikenUye",
         '=COUNTIFS(tblPaydas[Bakiye],">0",tblPaydas[Aktif],"EVET")', "0"),
        (12, "hbd_sureKuyruk",
         '=COUNTIF(tblAkis[Durum],"DURUSMA")+COUNTIF(tblAkis[Durum],"DAVADA")+COUNTIF(tblAkis[Durum],"VEKALET")', "0"),
        (13, "hbd_yetimKayit", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")', "0"),
        (14, "hbd_gecersizGecis", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', "0"),
        (15, "hbd_zincirKirik", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")', "0"),
        (16, "hbd_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))', YÜZDE),
        (17, "hbd_giderSapma", '=SUMIF(tblGider[GiderId],"<>",tblGider[Sapma])', TL),
        (18, "hbd_kararSayisi", '=COUNTA(tblKararlar[KararId])', "0"),
        (19, "hbd_kovaButunluk", "=ESLESTIRME!C2", None),
        (20, "hbd_kararKapi",
         '=IF(COUNTA(tblVekalet[BorcId])=0,"VERİ YOK",'
         'IF(OR(hbd_tahsilatOrani<hbd_hedefTahsilat*hbd_kritikCarpan,hbd_zincirKirik>0),"KRİTİK",'
         'IF(OR(hbd_tahsilatOrani<hbd_hedefTahsilat,hbd_yetimKayit>0,hbd_gecersizGecis>0),"DİKKAT",'
         '"UYGUN")))', None),
        (21, "hbd_kararGerekce",
         '=IF(hbd_kararKapi="VERİ YOK","Vekalet satırı yok — veri girin",'
         'IF(hbd_kararKapi="KRİTİK","Tahsilat oranı kritik veya ZİNCİR KIRIK var",'
         'IF(hbd_kararKapi="DİKKAT","Hedef altı tahsilat veya yetim/geçersiz geçiş",'
         '"Tahsilat ve akış eşikleri içinde")))', None),
        (22, "hbd_normAnahtar", '=COUNTA(tblAnahtarA[Norm])', "0"),
        (23, "hbd_tanimsizKova", '=COUNTIF(tblKok[NedenKodu],"TANIMSIZ")', "0"),
        (24, "hbd_durumHaritasi", '=COUNTA(AYARLAR!G6:G12)', "0"),
        (25, "hbd_kanitRaporu", '=IF(hbd_kararSayisi>0,1,)', "0"),
        (26, "hbd_kapaliDusum",
         '=COUNTIF(tblAkis[Durum],"KAPALI")+COUNTIF(tblAkis[Durum],"IPTAL")', "0"),
        (27, "hbd_modulT1", "=hbd_tahsilatOrani", YÜZDE),
        (28, "hbd_modulT1Yorum",
         '="Tahsilat oranı: "&TEXT(hbd_modulT1,"0.0%")', None),
        (29, "hbd_modulT2", "=hbd_trendAgustos", "0.00"),
        (30, "hbd_modulT2Yorum", '="Ağustos trend endeksi "&TEXT(hbd_modulT2,"0.00")', None),
        (31, "hbd_modulO1", "=hbd_gecikenUye", "0"),
        (32, "hbd_modulO1Yorum", '="Geciken aktif müvekkil: "&hbd_modulO1', None),
        (33, "hbd_modulO2",
         "=hbd_toplamBorc*hbd_gecikmeZammi", TL),
        (34, "hbd_modulO2Yorum", '="Tornado: gecikme zammı etkisi"', None),
        (35, "hbd_modulO3",
         "=hbd_toplamTahsilat*(hbd_senaryoIyimser-hbd_senaryoKotumser)", TL),
        (36, "hbd_modulO3Yorum", '="Senaryo bandı genişliği"', None),
        (37, "hbd_modulO6",
         '=IF(SUMIF(tblGider[GiderId],"<>",tblGider[Gerceklesen])=0,0,'
         'SUMIF(tblGider[Gerceklesen],">0")/100000)', "0.00"),
        (38, "hbd_modulO6Yorum", '="Gider yoğunlaşma vekili"', None),
        (39, "hbd_modulO8",
         '=IF(COUNTA(tblVekalet[BorcId])=0,0,'
         'ROUND(100*(1-(hbd_yetimKayit+hbd_gecersizGecis+hbd_zincirKirik)/10),0))', "0"),
        (40, "hbd_modulO8Yorum", '="Veri kalite skoru 0-100"', None),
        (41, "hbd_modulI1",
         "=hbd_toplamTahsilat*hbd_trendEylul", TL),
        (42, "hbd_modulI1Yorum", '="Sonraki ay tahmin (Eylül endeksi)"', None),
        (43, "hbd_modulI2",
         "=IFERROR(PERCENTILE(tblVekalet[GecikmeGun],0.9),0)", "0.0"),
        (44, "hbd_modulI2Yorum", '="Gecikme günü P90"', None),
        (45, "hbd_kvkkBeyani", 1, "0"),
        (46, "hbd_kritikCarpan", 0.85, "0.00"),
    ]
    for r, ad, formul, sayi in aynalar:
        h(ws, r, 13, ad)
        if isinstance(formul, str) and formul.startswith("="):
            h(ws, r, 14, formul, sayi=sayi)
        else:
            h(ws, r, 14, formul, sayi=sayi)

    genislik(ws, {"A": 22, "B": 14, "C": 10, "D": 36, "G": 14, "H": 14,
                  "I": 8, "J": 22, "K": 12, "M": 22, "N": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "GIRDI: Büro adı, dönem, asgari vekalet ve hedefleri sarı hücrelere yazın.",
        "PAYDASLAR: Her müvekkil için benzersiz MuvekkilId ve MuvekkilNo kullanın.",
        "DOSYA: Dosya kartını ve son süre tarihini girin; süre alarmı AYARLAR.hbd_sureEsikGun ile çalışır.",
        "VEKALET: Vekalet ücreti borçlandırma; gecikme zammı AYARLAR.hbd_gecikmeZammi ile hesaplanır.",
        "TAHSILAT: Vekalet tahsilatlarını müvekkil+dönem anahtarıyla girin; ANAHTAR/ESLESTIRME otomatik çalışır.",
        "AKIS: Dosya durumunu haritaya uygun ilerletin; süre ve geçersiz geçiş uyarılır.",
        "GIDER / KARARLAR: Büro masraf sapması ve dosya/vekalet kararlarını kaydedin.",
        "PANO / KARAR / RAPOR: Dosya/vekalet özeti ve müvekkil kanıt çıktısı.",
        "Korumalı alanlar şifre 1234 ile kilitlidir; sarı giriş hücreleri açıktır.",
        "Bu araç hukuki tavsiye değildir; AAÜT/HMK çerçevesinde genel kurul karar destek kanıtı üretir.",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, f"{i-4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
    genislik(ws, {"A": 90})


def ad_tanimlari(wb):
    param_bas = 5
    param_adlar = [
        "raporTarihi", "toleransKurus", "zincirTolerans",
        "hbd_gecikmeZammi", "hbd_sureEsikGun", "hbd_esikFark", "hbd_ikincilUzunluk",
        "hbd_hedefTahsilat", "hbd_senaryoIyimser", "hbd_senaryoKotumser",
        "hbd_fiyatModeli", "hbd_olcekHedef",
        "hbd_trendOcak", "hbd_trendSubat", "hbd_trendMart", "hbd_trendNisan",
        "hbd_trendMayis", "hbd_trendHaziran", "hbd_trendTemmuz", "hbd_trendAgustos",
        "hbd_trendEylul", "hbd_trendEkim", "hbd_trendKasim", "hbd_trendAralik",
    ]
    for i, ad in enumerate(param_adlar):
        ad_ekle(wb, ad, f"AYARLAR!$B${param_bas + i}")

    ayna = {
        "hbd_toplamBorc": 7, "hbd_toplamTahsilat": 8, "hbd_tahsilatOrani": 9,
        "hbd_kasaBakiye": 10, "hbd_gecikenUye": 11, "hbd_sureKuyruk": 12,
        "hbd_yetimKayit": 13, "hbd_gecersizGecis": 14, "hbd_zincirKirik": 15,
        "hbd_eslesmeOrani": 16, "hbd_giderSapma": 17, "hbd_kararSayisi": 18,
        "hbd_kovaButunluk": 19, "hbd_kararKapi": 20, "hbd_kararGerekce": 21,
        "hbd_normAnahtar": 22, "hbd_tanimsizKova": 23, "hbd_durumHaritasi": 24,
        "hbd_kanitRaporu": 25, "hbd_kapaliDusum": 26,
        "hbd_modulT1": 27, "hbd_modulT1Yorum": 28, "hbd_modulT2": 29,
        "hbd_modulT2Yorum": 30, "hbd_modulO1": 31, "hbd_modulO1Yorum": 32,
        "hbd_modulO2": 33, "hbd_modulO2Yorum": 34, "hbd_modulO3": 35,
        "hbd_modulO3Yorum": 36, "hbd_modulO6": 37, "hbd_modulO6Yorum": 38,
        "hbd_modulO8": 39, "hbd_modulO8Yorum": 40, "hbd_modulI1": 41,
        "hbd_modulI1Yorum": 42,         "hbd_modulI2": 43, "hbd_modulI2Yorum": 44,
        "hbd_kvkkBeyani": 45, "hbd_kritikCarpan": 46,
    }
    for ad, r in ayna.items():
        ad_ekle(wb, ad, f"AYARLAR!$N${r}")

    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$A$6:$A$7")
    ad_ekle(wb, "ListeDurum", "LISTELER!$C$6:$C$12")
    ad_ekle(wb, "ListeKararSonuc", "LISTELER!$E$6:$E$9")
    ad_ekle(wb, "ListeOrgan", "LISTELER!$G$6:$G$7")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$I$6:$I$20")


def kosullu_bicim(wb):
    kirmizi = Font(color="B3261E", bold=True)
    yesil = Font(color="1F7A4D", bold=True)
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    for sayfa, hucre in [("KARAR", "B5"), ("PANO", "B5"), ("HIZLI_BASLANGIC", "B10")]:
        ws = wb[sayfa]
        for metin, font, fill in [
            ('"VERİ YOK"', kirmizi, k_fill),
            ('"KRİTİK"', kirmizi, k_fill),
            ('"DİKKAT"', Font(color="B7791F", bold=True), a_fill),
            ('"UYGUN"', yesil, y_fill),
        ]:
            ws.conditional_formatting.add(hucre, FormulaRule(
                formula=[f'ISNUMBER(SEARCH({metin},{hucre}))'], font=font, fill=fill))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B9", CellIsRule(
        operator="lessThan", formula=["GIRDI!B9"], fill=a_fill))
    ws.conditional_formatting.add("B13", CellIsRule(
        operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B14", CellIsRule(
        operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B15", CellIsRule(
        operator="greaterThan", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B18", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("B18", CellIsRule(
        operator="greaterThanOrEqual", formula=["85"], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B11", CellIsRule(
        operator="greaterThan", formula=["3"], fill=a_fill))
    ws.conditional_formatting.add("B12", CellIsRule(
        operator="greaterThan", formula=["2"], fill=a_fill))
    ws.conditional_formatting.add("B16", CellIsRule(
        operator="lessThan", formula=["0.8"], fill=a_fill))
    ws.conditional_formatting.add("B17", CellIsRule(
        operator="greaterThan", formula=["5000"], font=kirmizi))

    ws = wb["KONTROLLER"]
    for txt, font, fill in [
        ("HATALI", kirmizi, k_fill), ("ALARM", kirmizi, k_fill),
        ("BOŞ", Font(color="B7791F"), a_fill), ("TEMİZ", yesil, y_fill),
        ("DOLU", yesil, y_fill),
    ]:
        ws.conditional_formatting.add("B5:B12", FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{txt}",B5))'], font=font, fill=fill))

    ws = wb["AKIS"]
    _son = min(ILK + 50, SON)
    ws.conditional_formatting.add(f"L{ILK}:L{_son}", FormulaRule(
        formula=['ISNUMBER(SEARCH("YETİM",L6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(f"M{ILK}:M{_son}", FormulaRule(
        formula=['ISNUMBER(SEARCH("GEÇERSİZ",M6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add(f"N{ILK}:N{_son}", FormulaRule(
        formula=['ISNUMBER(SEARCH("ZİNCİR",N6))'], font=kirmizi, fill=k_fill))

    ws = wb["ESLESTIRME"]
    ws.conditional_formatting.add("C2", FormulaRule(
        formula=['C2="KAYIP"'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("C2", FormulaRule(
        formula=['C2="BUTUN"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add(f"G{ILK}:G{min(ILK+40, SON)}", FormulaRule(
        formula=[f'G{ILK}="tutar_farki"'], fill=a_fill))
    ws.conditional_formatting.add(f"G{ILK}:G{min(ILK+40, SON)}", FormulaRule(
        formula=[f'G{ILK}="eslesmedi"'], font=kirmizi, fill=k_fill))

    ws = wb["GIDER"]
    ws.conditional_formatting.add(f"F{ILK}:F{min(ILK+30, SON)}", CellIsRule(
        operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add(f"F{ILK}:F{min(ILK+30, SON)}", CellIsRule(
        operator="lessThan", formula=["0"], font=yesil))

    ws = wb["PAYDASLAR"]
    ws.conditional_formatting.add(f"H{ILK}:H{min(ILK+30, SON)}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=a_fill))

    ws = wb["VEKALET"]
    ws.conditional_formatting.add(f"H{ILK}:H{min(ILK+30, SON)}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=a_fill))

    ws = wb["KARAR"]
    ws.conditional_formatting.add("B8", CellIsRule(
        operator="lessThan", formula=["B9"], fill=a_fill))
    ws.conditional_formatting.add("B10", CellIsRule(
        operator="greaterThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("KRİTİK",B6))'], font=kirmizi, fill=k_fill))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"D{ILK}:D{min(ILK+40, SON)}", FormulaRule(
        formula=['D6="TANIMSIZ"'], font=kirmizi, fill=k_fill))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B6:B12", CellIsRule(
        operator="lessThanOrEqual", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B9", CellIsRule(
        operator="lessThan", formula=["0.8"], fill=a_fill))


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (hizli_baslangic, "HIZLI_BASLANGIC"),
        (girdi, "GIRDI"),
        (paydaslar, "PAYDASLAR"),
        (dosya, "DOSYA"),
        (vekalet, "VEKALET"),
        (tahsilat, "TAHSILAT"),
        (anahtar, "ANAHTAR"),
        (eslestirme, "ESLESTIRME"),
        (kok_neden, "KOK_NEDEN"),
        (akis, "AKIS"),
        (gider, "GIDER"),
        (kararlar, "KARARLAR"),
        (pano, "PANO"),
        (karar, "KARAR"),
        (rapor, "RAPOR"),
        (kontroller, "KONTROLLER"),
        (ornek_veri, "ORNEK_VERI"),
        (listeler, "LISTELER"),
        (ayarlar, "AYARLAR"),
        (kilavuz, "KILAVUZ"),
    ]
    ilk = wb.active
    for i, (fn, _ad) in enumerate(siralar):
        ws = ilk if i == 0 else wb.create_sheet()
        fn(ws)

    ad_tanimlari(wb)
    kosullu_bicim(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=min(ILK + 80, SON))

    for ws in wb.worksheets:
        sayfa_koru(ws)

    # G15: başlık satırları taşmasın
    from openpyxl.styles import Alignment
    for sn in ("PAYDASLAR", "DOSYA", "VEKALET", "TAHSILAT", "ANAHTAR", "ESLESTIRME",
               "KOK_NEDEN", "AKIS", "GIDER", "KARARLAR"):
        if sn in wb.sheetnames:
            wb[sn].column_dimensions["A"].width = max(
                wb[sn].column_dimensions["A"].width or 14, 58)
            if wb[sn]["A3"].value:
                wb[sn]["A3"].alignment = Alignment(wrap_text=True, vertical="center")

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "HukukBurosuDosyaVekalet.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    hsh = hashlib.sha256()
    with open(hedef, "rb") as f:
        for b in iter(lambda: f.read(65536), b""):
            hsh.update(b)
    dig = hsh.hexdigest()

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        shutil.copy2(hedef, cikti)

    csv_yol = os.path.join(urun_dir, "ornek_veri.csv")
    with open(csv_yol, "w", encoding="utf-8") as f:
        f.write("VekaletId,MuvekkilId,Donem,BorcTutari,VadeTarihi\n")
        for bid, mid, don, tut, vade in VEKALET_SATIR:
            f.write(f"{bid},{mid},{don},{tut},{vade.isoformat()}\n")

    print(f"Dosya: {hedef}")
    print(f"Kopya: {cikti}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    main()
