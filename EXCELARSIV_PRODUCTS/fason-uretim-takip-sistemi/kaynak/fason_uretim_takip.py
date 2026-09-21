#!/usr/bin/env python3
"""Fason Üretim Takip Sistemi — A3 üretim betiği (manda v6)."""

from __future__ import annotations

import os
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
    CATI,
    FONT,
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
    GUN,
    KOYU_LACIVERT,
    NORMAL,
    ORTA_MAVI,
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
    sonraki_kimlik_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)

URUN_AD = "Fason Üretim Takip Sistemi"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "VERILDI", "durum_adi": "Verildi", "sira": 1,
     "izinli_sonraki_durumlar": "URETIMDE", "kategori": "acik"},
    {"durum_id": "URETIMDE", "durum_adi": "Üretimde", "sira": 2,
     "izinli_sonraki_durumlar": "TESLIM|IADE", "kategori": "acik"},
    {"durum_id": "TESLIM", "durum_adi": "Teslim", "sira": 3,
     "izinli_sonraki_durumlar": "KAPALI", "kategori": "acik"},
    {"durum_id": "IADE", "durum_adi": "İade", "sira": 4,
     "izinli_sonraki_durumlar": "KAPALI|URETIMDE", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "VERILDI|URETIMDE",
    "URETIMDE|TESLIM",
    "URETIMDE|IADE",
    "TESLIM|KAPALI",
    "IADE|KAPALI",
    "IADE|URETIMDE",
]

FASONCULAR = [
    ("FAS-00001", "Atlas Metal Fason", "Torna", "İstanbul"),
    ("FAS-00002", "Demir Kalıp A.Ş.", "Kalıp", "Ankara"),
    ("FAS-00003", "Nova Sac İşleme", "Sac", "İzmir"),
    ("FAS-00004", "Proje Kaplama", "Kaplama", "Bursa"),
    ("FAS-00005", "Su Enerji Montaj", "Montaj", "Antalya"),
    ("FAS-00006", "Cephe Boya Fason", "Boya", "Kocaeli"),
    ("FAS-00007", "Zemin Kaynak Ltd.", "Kaynak", "Adana"),
    ("FAS-00008", "İskele CNC", "CNC", "Gaziantep"),
]


def _yas_kova_fut(yas_hucre: str) -> str:
    """0-7 / 8-14 / 15-30 / 30+ — eşikler AYARLAR adlı aralıklarından."""
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=fut_kovaEsik1,"0-7",'
        f'IF({yas_hucre}<=fut_kovaEsik2,"8-14",'
        f'IF({yas_hucre}<=fut_kovaEsik3,"15-30","30+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Fason kuyruğu ve kararı etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["VERILDI", "URETIMDE", "TESLIM", "IADE", "KAPALI",
                "URETIMDE", "TESLIM", "VERILDI"]
    oncekiler = ["", "VERILDI", "URETIMDE", "URETIMDE", "TESLIM",
                 "VERILDI", "URETIMDE", "VERILDI"]
    for i in range(1, DEMO_AKIS + 1):
        kart = FASONCULAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "FAS-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: VERILDI → TESLIM
            o, d = "VERILDI", "TESLIM"
        if i == 20:  # geçersiz: TESLIM → URETIMDE
            o, d = "TESLIM", "URETIMDE"
        verilen = 1000 + i * 50
        if d in ("TESLIM", "KAPALI"):
            fire = 20 + (i % 15)
            donen = verilen - fire
        elif d == "IADE":
            fire = 40 + (i % 10)
            donen = max(0, verilen - fire - 100)
        elif d == "URETIMDE":
            fire = 0
            donen = 0
        else:
            fire = 0
            donen = 0
        if i in (12, 19):  # zincir kırık: dönen+fire ≠ verilen
            donen = verilen + 80
            fire = 25
        acik = 0 if d == "KAPALI" else max(0, verilen - donen - fire)
        if i in (12, 19) and d != "KAPALI":
            acik = max(0, verilen)  # sapma görünür; açık bakiyeyi tut
            # Zincir için dönen+fire verileni aşacak
        maliyet = 45 + (i % 20)
        verilis = RAPOR_TARIHI - timedelta(days=(i * 3) % 45)
        plan_teslim = verilis + timedelta(days=10 + (i % 8))
        satirlar.append({
            "id": f"FEM-{i:05d}",
            "kart": kart,
            "siparis": f"SIP-2026-{i:04d}",
            "verilis": verilis,
            "plan_teslim": plan_teslim,
            "onceki": o,
            "durum": d,
            "verilen": verilen,
            "donen": donen,
            "fire": fire,
            "tutar": acik,
            "maliyet": maliyet,
            "is_turu": FASONCULAR[(i - 1) % DEMO_KART][2],
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Fason notu {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Fason iş emirlerini verilen↔dönen+fire zinciri, durum makinesi ve fire/gecikme "
      "eşikleri ile yönetir; açık bakiyeyi görünür kılar; DEVAM / UYAR / DURDUR kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Verildi → Üretimde → Teslim / İade → Kapalı",
        "Verilen ↔ dönen + fire zinciri (ZİNCİR KIRIK bayrağı)",
        "Açık bakiye + fire oranı + gecikme kuyruğu",
        "Geçersiz geçiş ve yetim fasoncu uyarıları",
        "Dönemsel fason kanıt raporu — yönetici/SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Üretim planlama / fason takip — açık bakiye ve fire kuyruğu",
        "Fabrika müdürü — DEVAM / UYAR / DURDUR kararı",
        "SMMM / mali işler — fason kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a fasoncu kartı girin. 2) AKIS'e iş emri satırlarını yazın. "
      "3) KUYRUKLAR ve PANO'dan günün kararını görün. 4) RAPOR'dan kanıt çıktısı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblAkis[IslemId])=0,"VERİ YOK",'
      'IF(OR(fut_yetimKayit>0,fut_gecersizGecis>0,'
      'fut_fireOran>fut_esikFireKritik,fut_gecikenAdet>fut_esikGecikenKritik),"DURDUR",'
      'IF(OR(fut_fireOran>fut_esikFireUyari,fut_gecikenAdet>fut_esikGecikenUyari,'
      'fut_zincirKirik>0),"UYAR","DEVAM")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: DEVAM | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Bakiye", "=fut_acikBakiye", CATI),
        (7, "Fire Oranı Ort.", "=fut_fireOran", YÜZDE),
        (8, "Verildi Adet", '=COUNTIFS(tblAkis[Durum],"VERILDI",tblAkis[Tutar],">0")', CATI),
        (9, "Üretimde Adet", '=COUNTIFS(tblAkis[Durum],"URETIMDE",tblAkis[Tutar],">0")', CATI),
        (10, "Geciken Adet", "=fut_gecikenAdet", CATI),
        (11, "Fire Sapma", "=fut_fireSapma", YÜZDE),
        (12, "Yetim Kayıt", "=fut_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=fut_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=fut_zincirKirik", CATI),
        (15, "Tornado Zirve", "=fut_tornadoZirve", TL),
        (16, "Senaryo Farkı", "=fut_senaryoKarsilastirma", TL),
        (17, "Tahmin Aralık", "=fut_tahminAralik", CATI),
        (18, "HHI Yoğunlaşma", "=fut_hhi", YÜZDE),
        (19, "Kalite Skoru", "=fut_kaliteSkor", CATI),
        (20, "P90 Yüzdelik", "=fut_yuzdelikP90", CATI),
        (21, "Kapalı Düşüm", "=fut_kapaliDusum", CATI),
        (22, "Senaryo Motor", "=fut_senaryoMotor", None),
        (23, "Karar Metni", "=fut_kararMetni", None),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=fut_yorumBakiye"),
        (8, "=fut_yorumFire"),
        (9, "=fut_yorumSapma"),
        (10, "=fut_yorumTornado"),
        (11, "=fut_yorumSenaryo"),
        (12, "=fut_yorumHhi"),
        (13, "=fut_yorumKalite"),
        (14, "=fut_yorumSenMotor"),
        (15, "=fut_yorumTahmin"),
        (16, "=fut_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Verildi", 6), ("Durum Üretimde", 8), ("Durum Teslim", 5),
        ("Durum İade", 5), ("Durum Kapalı", 4),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Yaş Kova")
    h(ws, 26, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 7), ("8-14", 8), ("15-30", 6), ("30+", 7),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Teslim")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 1200 + i * 80, sayi=CATI)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 380000), ("Baz", 450000), ("İyimser", 520000),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=31), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=31))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık Teslim"
    c3.add_data(Reference(ws, min_col=8, min_row=26, max_row=34), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=27, max_row=34))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A53")

    c4 = BarChart()
    c4.title = "Senaryo Karşılaştırma"
    c4.add_data(Reference(ws, min_col=11, min_row=26, max_row=29), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=27, max_row=29))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F53")

    h(ws, 26, 13, "Uyarı")
    h(ws, 26, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 9),
    ], 27):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 26, 16, "Ay")
    h(ws, 26, 17, "Fire")
    for i in range(6):
        h(ws, 27 + i, 16, i + 1, sayi=CATI)
        h(ws, 27 + i, 17, 40 + i * 8, sayi=CATI)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = LineChart()
    c6.title = "Aylık Fire"
    c6.add_data(Reference(ws, min_col=17, min_row=26, max_row=32), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=27, max_row=32))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F68")

    c7 = PieChart()
    c7.title = "Kova Payı"
    c7.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c7.width, c7.height = 9, 7
    ws.add_chart(c7, "A83")

    c8 = BarChart()
    c8.title = "Durum Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=26, max_row=31), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=27, max_row=31))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F83")

    baski_hazirla(ws, "A1:N36", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Fasoncu kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=fut_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: FAS-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "IsTuru", "Sehir", "KapasiteGun",
                "RiskNotu", "Aktif", "ToplamAcik", "OrtGecikme"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamAcik": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtGecikme": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[GecikmeGun]),0))'
        ),
    }
    for i, (kid, unvan, is_turu, sehir) in enumerate(FASONCULAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="FAS-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Fasoncu unvanı")
        _sari(ws, r, 3, is_turu, baslik="İş Türü", mesaj="İş türü")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, 800 + i * 50, sayi=CATI, baslik="Kapasite", mesaj="Günlük kapasite adet")
        _sari(ws, r, 6, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=CATI if c == 5 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="FAS-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "İş Türü"), (4, "Şehir"), (6, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "whole", "0", f"E{ILK}:E{SON}",
              baslik="Kapasite", mesaj="Kapasite ≥ 0",
              hata_baslik="Kapasite", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 22


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=24)
    alt_bant(ws, 3,
             "Fason iş emri satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "SiparisNo", "VerilisTarihi", "PlanTeslim",
        "OncekiDurum", "Durum", "VerilenMiktar", "DonenMiktar", "FireMiktar",
        "Tutar", "BirimMaliyet", "IsTuru", "YorumA", "YorumB",
        "ToplamDonus", "FireOranSatir", "GecikmeGun", "YasKova",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu", "MaliyetTl",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[PlanTeslim]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_fut("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[VerilenMiktar]]",
        "tblAkis[[#This Row],[ToplamDonus]]",
        "zincirTolerans",
    ).lstrip("=")
    zincir_guvenli = (
        f'IF(OR(tblAkis[[#This Row],[Durum]]="TESLIM",'
        f'tblAkis[[#This Row],[Durum]]="KAPALI",'
        f'tblAkis[[#This Row],[Durum]]="IADE",'
        f'tblAkis[[#This Row],[ToplamDonus]]>0),{zincir_f},"OK")'
    )

    form = {
        "ToplamDonus": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[DonenMiktar]]+tblAkis[[#This Row],[FireMiktar]])'
        ),
        "FireOranSatir": (
            'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            'tblAkis[[#This Row],[VerilenMiktar]]=0),"",'
            'tblAkis[[#This Row],[FireMiktar]]/tblAkis[[#This Row],[VerilenMiktar]])'
        ),
        "GecikmeGun": (
            f'IF(tblAkis[[#This Row],[IslemId]]="","",'
            f'IF(OR(tblAkis[[#This Row],[Durum]]="KAPALI",'
            f'tblAkis[[#This Row],[Durum]]="TESLIM"),0,{yas_f}))'
        ),
        "YasKova": f'IF(tblAkis[[#This Row],[IslemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_guvenli})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
        "MaliyetTl": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[Tutar]]*tblAkis[[#This Row],[BirimMaliyet]])'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="FEM-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["siparis"], baslik="Sipariş No", mesaj="Sipariş numarası")
        _sari(ws, r, 4, s["verilis"], sayi=TARİH, baslik="Veriliş Tarihi", mesaj="Veriliş tarihi")
        _sari(ws, r, 5, s["plan_teslim"], sayi=TARİH, baslik="Plan Teslim",
              mesaj="Planlanan teslim tarihi")
        _sari(ws, r, 6, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 7, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 8, s["verilen"], sayi=CATI, baslik="Verilen Miktar", mesaj="Verilen adet")
        _sari(ws, r, 9, s["donen"], sayi=CATI, baslik="Dönen Miktar", mesaj="Dönen adet")
        _sari(ws, r, 10, s["fire"], sayi=CATI, baslik="Fire Miktar", mesaj="Fire adet")
        _sari(ws, r, 11, s["tutar"], sayi=CATI, baslik="Açık Bakiye", mesaj="Açık bakiye adet")
        _sari(ws, r, 12, s["maliyet"], sayi=TL, baslik="Birim Maliyet", mesaj="Birim maliyet TL")
        _sari(ws, r, 13, s["is_turu"], baslik="İş Türü", mesaj="İş türü")
        _sari(ws, r, 14, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 15, s["yorum_b"], baslik="Yorum B", mesaj="Serbest takip notu")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 16):
            fmt = TARİH if c in (4, 5) else (
                TL if c == 12 else (CATI if c in (8, 9, 10, 11) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:X{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="FEM-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Sipariş No", mesaj="Sipariş numarası",
              hata_baslik="Sipariş", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Veriliş", mesaj="Veriliş tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "date", "1", f"E{ILK}:E{SON}",
              baslik="Plan Teslim", mesaj="Planlanan teslim",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"G{ILK}:G{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(8, "Verilen"), (9, "Dönen"), (10, "Fire"), (11, "Açık")]:
        dogrulama(ws, "whole", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"L{ILK}:L{SON}",
              baslik="Birim Maliyet", mesaj="≥ 0 TL",
              hata_baslik="Maliyet", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeYorumAB", f"N{ILK}:N{SON}",
              baslik="Yorum A", mesaj="A veya B",
              hata_baslik="Liste", hata_mesaj="A veya B")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 25)})
    ws.column_dimensions["O"].width = 16
    ws.column_dimensions["U"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Bakiye", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Kuyruk Verildi")
    h(ws, 6, 2,
      '=COUNTIFS(tblAkis[Durum],"VERILDI",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 6, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"VERILDI")', sayi=CATI)

    h(ws, 7, 1, "Kuyruk Üretimde")
    h(ws, 7, 2,
      '=COUNTIFS(tblAkis[Durum],"URETIMDE",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 7, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"URETIMDE")', sayi=CATI)

    h(ws, 8, 1, "Kuyruk Teslim")
    h(ws, 8, 2,
      '=COUNTIFS(tblAkis[Durum],"TESLIM",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 8, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"TESLIM")', sayi=CATI)

    h(ws, 9, 1, "Kuyruk İade")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[Durum],"IADE",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"IADE")', sayi=CATI)

    h(ws, 10, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 10, 2,
      '=COUNTIFS(tblAkis[GecikmeGun],">"&fut_esikGecikmeGun,tblAkis[Durum],"<>KAPALI",'
      'tblAkis[Durum],"<>TESLIM",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 10, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&fut_esikGecikmeGun,'
      'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>TESLIM")',
      sayi=CATI)

    h(ws, 12, 1, "Kapalı — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 1, "Kapalı")
    h(ws, 13, 2, '=COUNTIF(tblAkis[Durum],"KAPALI")', sayi=CATI)
    h(ws, 13, 3, '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[Tutar])', sayi=CATI)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=fut_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-14", "15-30", "30+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=CATI)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam işlem", '=COUNTA(tblAkis[IslemId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Açık bakiye toplam", "=SUM(tblAkis[Tutar])"),
        ("H04", "Verilen toplam", "=SUM(tblAkis[VerilenMiktar])"),
        ("H05", "Dönen toplam", "=SUM(tblAkis[DonenMiktar])"),
        ("H06", "Fire toplam", "=SUM(tblAkis[FireMiktar])"),
        ("H07", "Verildi adet", '=COUNTIF(tblAkis[Durum],"VERILDI")'),
        ("H08", "Üretimde adet", '=COUNTIF(tblAkis[Durum],"URETIMDE")'),
        ("H09", "Teslim adet", '=COUNTIF(tblAkis[Durum],"TESLIM")'),
        ("H10", "İade adet", '=COUNTIF(tblAkis[Durum],"IADE")'),
        ("H11", "Kapalı adet", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        ("H12", "Açık kuyruk adet", "=C12+C13+C14+C15"),
        ("H13", "Kapalı düşüm adet", "=C16"),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H19", "Yaş 8-14", '=COUNTIF(tblAkis[YasKova],"8-14")'),
        ("H20", "Yaş 15-30", '=COUNTIF(tblAkis[YasKova],"15-30")'),
        ("H21", "Yaş 30+", '=COUNTIF(tblAkis[YasKova],"30+")'),
        ("H22", "Ort gecikme gün",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[GecikmeGun]),0)"),
        ("H23", "Geciken adet",
         '=COUNTIFS(tblAkis[GecikmeGun],">"&fut_esikGecikmeGun,tblAkis[Durum],"<>KAPALI",'
         'tblAkis[Durum],"<>TESLIM",tblAkis[Tutar],">0")'),
        ("H24", "Geciken bakiye",
         '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&fut_esikGecikmeGun,'
         'tblAkis[Durum],"<>KAPALI",tblAkis[Durum],"<>TESLIM")'),
        ("H25", "Fire oranı ort",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[FireOranSatir]),0)"),
        ("H26", "Hedef fire sapma", "=C30-fut_hedefFire"),
        ("H27", "Açık maliyet TL",
         "=IFERROR(SUMIF(tblAkis[KayitDolu],1,tblAkis[MaliyetTl]),0)"),
        ("H28", "Temkinli senaryo", "=C32*fut_senaryoTemkinli"),
        ("H29", "Baz senaryo", "=C32*fut_senaryoBaz"),
        ("H30", "İyimser senaryo", "=C32*fut_senaryoIyimser"),
        ("H31", "Senaryo farkı", "=C35-C33"),
        ("H32", "Tornado (oran×maliyet)", "=C32*fut_tornadoOran"),
        ("H33", "HHI proxy",
         "=IFERROR((MAX(C12,C13,C14,C15)/MAX(C17,1))^2,0)"),
        ("H34", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,100-C19*10-C20*10-C21*10-C28*2))"),
        ("H35", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],fut_yuzdelikOran),0)"),
        ("H36", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C40)"),
        ("H37", "Tahmin aralık", "=IFERROR((C40+C41)/2,0)"),
        ("H38", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[VerilenMiktar],fut_yuzdelikOran),0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Kapasite toplam", "=SUM(tblKartlar[KapasiteGun])"),
        ("H41", "Kapasite kullanım", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Verildi",C12,"Üretimde",C13,"Teslim",C14,'
         '"İade",C15,"Geciken",C28)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C23,C24,C25,C26)'),
        ("H44", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C33,"Baz",C34,"İyimser",C35)'),
        ("H45", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H46", "Fire kritik bayrak",
         '=IF(C30>fut_esikFireKritik,1,0)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "maliyet", "senaryo", "tornado", "fark")) and "oran" not in acik.lower() and "adet" not in acik.lower() and "metin" not in acik.lower() and "özet" not in acik.lower() and "bayrak" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "hhi" in acik.lower() or "hazırlık" in acik.lower() or "sapma" in acik.lower() or "kullanım" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "gün" in acik.lower():
            fmt = GUN
        if "kalite" in acik.lower() or "bayrak" in acik.lower():
            fmt = CATI
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 55, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 56, 1, "Temkinli")
    h(ws, 56, 2, "=fut_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 56, 3, "=C33", sayi=TL)
    h(ws, 57, 1, "Baz")
    h(ws, 57, 2, "=fut_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 57, 3, "=C34", sayi=TL)
    h(ws, 58, 1, "İyimser")
    h(ws, 58, 2, "=fut_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 58, 3, "=C35", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Fason Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Bakiye", "=fut_acikBakiye", CATI),
        (10, "Fire Oranı", "=fut_fireOran", YÜZDE),
        (11, "Fire Sapma", "=fut_fireSapma", YÜZDE),
        (12, "Yetim", "=fut_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=fut_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=fut_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=fut_tahminAralik", CATI),
        (16, "Kanıt Özeti", "=fut_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Durum haritası AYARLAR'dadır (FUT-001). Geçersiz geçiş engellenmez, görünür kılınır "
      "(FUT-002). Verilen = dönen + fire zinciri toleransla denetlenir (FUT-003). "
      "Kapalı kayıtlar kuyruktan düşer (FUT-004). Fire eşikleri kararı belirler (FUT-005). "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 23, 1, "2. Fire eşiğini aşan fasoncuları durdurun veya yeniden müzakere edin.")
    h(ws, 24, 1, "3. Zincir kırık satırlarda verilen/dönen/fire farkını açıklayın.")
    h(ws, 26, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H26", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Verildi", '=COUNTIF(tblAkis[Durum],"VERILDI")'),
        (8, "Üretimde", '=COUNTIF(tblAkis[Durum],"URETIMDE")'),
        (9, "Teslim", '=COUNTIF(tblAkis[Durum],"TESLIM")'),
        (10, "İade", '=COUNTIF(tblAkis[Durum],"IADE")'),
        (11, "Kapalı", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        (12, "Açık toplam", "=B7+B8+B9+B10"),
        (13, "Yetim", "=fut_yetimKayit"),
        (14, "Geçersiz", "=fut_gecersizGecis"),
        (15, "Zincir kırık", "=fut_zincirKirik"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=fut_kaliteSkor"),
        (18, "Açık bakiye", "=fut_acikBakiye"),
        (19, "Fire oranı", "=fut_fireOran"),
        (20, "Kapalı düşüm", "=fut_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=YÜZDE if r == 19 else (CATI if r != 16 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} fasoncu kartı ve {DEMO_AKIS} iş emri satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Verildi/Üretimde/Teslim/İade/Kapalı — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 20000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "DurumId", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUM_HARITASI, 7):
        _sari(ws, i, 1, d["durum_id"], baslik="Durum", mesaj="Durum kimliği")

    h(ws, 3, 3, "İzinli Geçişler", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Gecis", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, g in enumerate(IZINLI_GECIS, 7):
        _sari(ws, i, 3, g, baslik="Geçiş", mesaj="eski|yeni")

    h(ws, 3, 5, "Evet / Hayır", kalin=True)
    h(ws, 6, 5, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 5, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 5, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")

    h(ws, 3, 7, "Yorum A/B", kalin=True)
    h(ws, 6, 7, "Yorum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 7, "A", baslik="Yorum", mesaj="A veya B")
    _sari(ws, 8, 7, "B", baslik="Yorum", mesaj="A veya B")

    h(ws, 3, 9, "İş Türü", kalin=True)
    h(ws, 6, 9, "IsTuru", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Torna", "Kalıp", "Sac", "Kaplama", "Montaj", "Boya", "Kaynak", "CNC"], 7):
        _sari(ws, i, 9, b, baslik="İş Türü", mesaj="İş türü listesi")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası + FUT kuralları", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 5, "adet", "İş kuralı — verilen/(dönen+fire) sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("fut_esikGecikmeGun", 7, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("fut_esikGecikenUyari", 3, "adet", "İç politika", "01.01.2026", "UYAR geciken eşiği"),
        ("fut_esikGecikenKritik", 8, "adet", "İç politika", "01.01.2026", "DURDUR geciken eşiği"),
        ("fut_kovaEsik1", 7, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("fut_kovaEsik2", 14, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("fut_kovaEsik3", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("fut_hedefFire", 0.03, "oran", "Fire modeli", "01.01.2026", "Hedef fire oranı"),
        ("fut_esikFireUyari", 0.05, "oran", "Fire modeli FUT-005", "01.01.2026", "UYAR fire eşiği"),
        ("fut_esikFireKritik", 0.10, "oran", "Fire modeli FUT-005", "01.01.2026", "DURDUR fire eşiği"),
        ("fut_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("fut_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("fut_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("fut_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("fut_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("fut_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("fut_azamiMiktar", 500000, "adet", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("fut_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi FUT-001",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("fut_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("fut_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06 FUT-004",
         "01.01.2026", "Kapalı düşüm kuralı"),
        ("fut_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("fut_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("fut_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("fut_kuralFUT001", "FUT-001", "kural", "Durum haritası zorunlu", "01.01.2026",
         "Durum makinesi AYARLAR'da"),
        ("fut_kuralFUT002", "FUT-002", "kural", "Geçersiz geçiş görünür", "01.01.2026",
         "Engellemez, bayrak üretir"),
        ("fut_kuralFUT003", "FUT-003", "kural", "Verilen=dönen+fire zinciri", "01.01.2026",
         "ZİNCİR KIRIK"),
        ("fut_kuralFUT004", "FUT-004", "kural", "Kapalı kuyruktan düşer", "01.01.2026",
         "A06 kapalı düşüm"),
        ("fut_kuralFUT005", "FUT-005", "kural", "Fire eşik kararları", "01.01.2026",
         "UYAR/DURDUR eşikleri"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.15:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 37, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 38, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 39 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H38:L43", dh_bas)

    h(ws, 46, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 47, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 47, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (48, "fut_acikBakiye", "=HESAP!C8"),
        (49, "fut_fireOran", "=HESAP!C30"),
        (50, "fut_gecikenAdet", "=HESAP!C28"),
        (51, "fut_fireSapma", "=HESAP!C31"),
        (52, "fut_yetimKayit", "=HESAP!C19"),
        (53, "fut_gecersizGecis", "=HESAP!C20"),
        (54, "fut_zincirKirik", "=HESAP!C21"),
        (55, "fut_tahminAralik", "=HESAP!C42"),
        (56, "fut_senaryoKarsilastirma", "=HESAP!C36"),
        (57, "fut_tornadoZirve", "=HESAP!C37"),
        (58, "fut_hhi", "=HESAP!C38"),
        (59, "fut_kaliteSkor", "=HESAP!C39"),
        (60, "fut_yuzdelikP90", "=HESAP!C43"),
        (61, "fut_senaryoMotor", "=HESAP!C49"),
        (62, "fut_kapaliDusum", "=HESAP!C18"),
        (63, "fut_kuyrukOzeti", "=HESAP!C47"),
        (64, "fut_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C8,"Fire%",TEXT(HESAP!C30,"0.0%"),'
         '"Yetim",HESAP!C19,"Zincir",HESAP!C21,"Geciken",HESAP!C28)'),
        (65, "fut_sonrakiKimlik",
         sonraki_kimlik_formul("FAS", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
        (66, "fut_kararMetni",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",PANO!B4,"— açık bakiye",fut_acikBakiye,'
         '"fire",TEXT(fut_fireOran,"0.0%"))'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 68, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (69, "fut_yorumBakiye",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Açık bakiye",fut_acikBakiye,"adet")'),
        (70, "fut_yorumFire",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ort fire oranı",TEXT(fut_fireOran,"0.0%"))'),
        (71, "fut_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Fire sapma",TEXT(fut_fireSapma,"0.0%"))'),
        (72, "fut_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(fut_tornadoZirve,"0.00"),"TL")'),
        (73, "fut_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(fut_senaryoKarsilastirma,"0.00"),"TL")'),
        (74, "fut_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(fut_hhi,"0.0%"))'),
        (75, "fut_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",fut_kaliteSkor)'),
        (76, "fut_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",fut_senaryoMotor)'),
        (77, "fut_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",fut_tahminAralik)'),
        (78, "fut_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 verilen",fut_yuzdelikP90)'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına fasoncu kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her fason iş emrini yazın; durum kolonunu listeden seçin.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — uyarı üretir (makrosuz kural, FUT-002).",
        "5. Kapalı durum kuyruklardan düşer (FUT-004).",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma plan teslime göredir.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 35):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 12):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Durum", mesaj="Durum kimliği",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    for r in range(7, 13):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Geçiş", mesaj="eski|yeni",
                  hata_baslik="Geçiş", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="40")
    ws = wb["RAPOR"]
    for r, ad in [(5, "Sürüm"), (22, "Aksiyon1"), (23, "Aksiyon2"), (24, "Aksiyon3")]:
        dogrulama(ws, "textLength", "0", f"B{r}" if r == 5 else f"A{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KILAVUZ"]
    for r in range(5, 14):
        dogrulama(ws, "textLength", "0", f"A{r}",
                  baslik="Kılavuz", mesaj="Bilgi satırı",
                  hata_baslik="Metin", hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"DEVAM"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"UYAR"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 24):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["AKIS"]
    for col, val, fill, font in [
        ("T", '"YETİM"', kirmizi, kf),
        ("U", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("V", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("V", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("URETIMDE", sari), ("IADE", kirmizi), ("TESLIM", yesil)]:
        ws.conditional_formatting.add(
            f"G{ILK}:G{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 11):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
    for r in range(18, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["HESAP"]
    for r in range(6, 52):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="equal", formula=['"VERİ VAR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B17", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B17", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))
    for r in range(5, 16):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"DEVAM"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"UYAR"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(9, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KARTLAR"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["1000"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans",
        "fut_esikGecikmeGun", "fut_esikGecikenUyari", "fut_esikGecikenKritik",
        "fut_kovaEsik1", "fut_kovaEsik2", "fut_kovaEsik3",
        "fut_hedefFire", "fut_esikFireUyari", "fut_esikFireKritik",
        "fut_senaryoTemkinli", "fut_senaryoBaz", "fut_senaryoIyimser",
        "fut_tornadoOran", "fut_yuzdelikOran", "fut_olcekHedef", "fut_azamiMiktar",
        "fut_durumHaritasi", "fut_kuyrukOzeti_metin", "fut_kapaliDusum_metin",
        "fut_kanitRaporu_metin", "fut_dosyaSurumu", "fut_firmaUnvan",
        "fut_kuralFUT001", "fut_kuralFUT002", "fut_kuralFUT003",
        "fut_kuralFUT004", "fut_kuralFUT005",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "fut_acikBakiye": 48, "fut_fireOran": 49, "fut_gecikenAdet": 50,
        "fut_fireSapma": 51, "fut_yetimKayit": 52, "fut_gecersizGecis": 53,
        "fut_zincirKirik": 54, "fut_tahminAralik": 55,
        "fut_senaryoKarsilastirma": 56, "fut_tornadoZirve": 57, "fut_hhi": 58,
        "fut_kaliteSkor": 59, "fut_yuzdelikP90": 60, "fut_senaryoMotor": 61,
        "fut_kapaliDusum": 62, "fut_kuyrukOzeti": 63, "fut_kanitRaporu": 64,
        "fut_sonrakiKimlik": 65, "fut_kararMetni": 66,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "fut_yorumBakiye": 69, "fut_yorumFire": 70, "fut_yorumSapma": 71,
        "fut_yorumTornado": 72, "fut_yorumSenaryo": 73, "fut_yorumHhi": 74,
        "fut_yorumKalite": 75, "fut_yorumSenMotor": 76, "fut_yorumTahmin": 77,
        "fut_yorumP90": 78,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$G$7:$G$8")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (kartlar, "KARTLAR"),
        (akis, "AKIS"),
        (kuyruklar, "KUYRUKLAR"),
        (hesap, "HESAP"),
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

    adlari_bagla(wb)
    _cok_dogrulama(wb)
    kosullu_bicim(wb)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=SON)

    for ws in wb.worksheets:
        sayfa_koru(ws)

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "FasonUretimTakipSistemi.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti), exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
