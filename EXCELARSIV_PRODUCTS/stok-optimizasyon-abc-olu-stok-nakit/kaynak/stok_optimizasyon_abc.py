#!/usr/bin/env python3
"""Stok Optimizasyon + ABC + Ölü Stok Nakit — A3 üretim betiği (manda v6)."""

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

URUN_AD = "Stok Optimizasyon + ABC + Ölü Stok Nakit"
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
    {"durum_id": "SAYIM", "durum_adi": "Sayım", "sira": 1,
     "izinli_sonraki_durumlar": "ABC", "kategori": "acik"},
    {"durum_id": "ABC", "durum_adi": "ABC", "sira": 2,
     "izinli_sonraki_durumlar": "AKSIYON|SAYIM", "kategori": "acik"},
    {"durum_id": "AKSIYON", "durum_adi": "Aksiyon", "sira": 3,
     "izinli_sonraki_durumlar": "KAPALI|ABC", "kategori": "acik"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "SAYIM|ABC",
    "ABC|AKSIYON",
    "ABC|SAYIM",
    "AKSIYON|KAPALI",
    "AKSIYON|ABC",
]

SKU_KARTLARI = [
    ("SKU-00001", "Çelik Sac 2mm", "Hammadde", "Merkez"),
    ("SKU-00002", "Motor Rulman 6205", "Yedek Parça", "Merkez"),
    ("SKU-00003", "Ambalaj Karton A", "Ambalaj", "Yan Depo"),
    ("SKU-00004", "Vida M8 Set", "Sarf", "Merkez"),
    ("SKU-00005", "Hidrolik Yağ 20L", "Sarf", "Yan Depo"),
    ("SKU-00006", "PLC Modül S7", "Yedek Parça", "Kritik"),
    ("SKU-00007", "Kayış Profil B", "Yedek Parça", "Merkez"),
    ("SKU-00008", "Etiket Rulo", "Ambalaj", "Yan Depo"),
]


def _yas_kova_soa(yas_hucre: str) -> str:
    """0-7 / 8-30 / 31-90 / 90+ — eşikler AYARLAR adlı aralıklarından."""
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=soa_kovaEsik1,"0-7",'
        f'IF({yas_hucre}<=soa_kovaEsik2,"8-30",'
        f'IF({yas_hucre}<=soa_kovaEsik3,"31-90","90+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Stok ABC ve ölü stok nakit kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["SAYIM", "ABC", "AKSIYON", "KAPALI",
                "ABC", "AKSIYON", "SAYIM", "ABC"]
    oncekiler = ["", "SAYIM", "ABC", "AKSIYON",
                 "SAYIM", "ABC", "SAYIM", "SAYIM"]
    abc_sinif = ["A", "B", "C", "A", "B", "C", "A", "B"]
    for i in range(1, DEMO_AKIS + 1):
        kart = SKU_KARTLARI[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "SKU-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: SAYIM → AKSIYON
            o, d = "SAYIM", "AKSIYON"
        if i == 20:  # geçersiz: KAPALI → SAYIM (önceki KAPALI)
            o, d = "KAPALI", "SAYIM"
        stok = 200 + i * 40
        maliyet = 25 + (i % 30)
        hesap_nakit = stok * maliyet
        kontrol = hesap_nakit
        if i in (12, 19):  # zincir kırık
            kontrol = hesap_nakit + 850
        stok_gunu = 10 + (i * 7) % 120
        if i in (8, 15, 24):
            stok_gunu = 110 + i  # ölü stok
        onerilen = 0 if stok_gunu > 90 else max(0, 80 - (i % 50))
        if d == "KAPALI":
            onerilen = 0
        son_hareket = RAPOR_TARIHI - timedelta(days=stok_gunu)
        sayim = RAPOR_TARIHI - timedelta(days=(i * 2) % 20)
        satirlar.append({
            "id": f"SOA-{i:05d}",
            "kart": kart,
            "sku": f"STK-2026-{i:04d}",
            "sayim": sayim,
            "son_hareket": son_hareket,
            "onceki": o,
            "durum": d,
            "stok": stok,
            "maliyet": maliyet,
            "kontrol": kontrol,
            "onerilen": onerilen,
            "stok_gunu": stok_gunu,
            "abc": abc_sinif[(i - 1) % len(abc_sinif)],
            "sinif": SKU_KARTLARI[(i - 1) % DEMO_KART][2],
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Stok notu {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Stok sayımını ABC sınıfı ve ölü stok nakit zinciri ile yönetir; "
      "önerilen siparişi görünür kılar; SİPARİŞ / BEKLE / ÖLÜ STOK ERİT kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Sayım → ABC → Aksiyon → Kapalı",
        "Stok adet × birim maliyet ↔ kontrol nakit zinciri (ZİNCİR KIRIK bayrağı)",
        "Ölü stok nakit + stok günü + önerilen sipariş kuyruğu",
        "Geçersiz geçiş ve yetim SKU uyarıları",
        "Dönemsel stok kanıt raporu — yönetici/SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Satın alma / depo planlama — ABC sınıfı ve önerilen sipariş",
        "Fabrika müdürü — SİPARİŞ / BEKLE / ÖLÜ STOK ERİT kararı",
        "SMMM / mali işler — ölü stok nakit kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a SKU kartı girin. 2) AKIS'e sayım/ABC satırlarını yazın. "
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
      'IF(OR(soa_oluStokNakit>soa_esikOluNakit,soa_oluAdet>soa_esikOluAdet),"ÖLÜ STOK ERİT",'
      'IF(soa_onerilenSiparis>soa_esikSiparis,"SİPARİŞ","BEKLE")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: BEKLE | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Ölü Stok Nakit", "=soa_oluStokNakit", TL),
        (7, "Önerilen Sipariş Ort.", "=soa_onerilenSiparis", CATI),
        (8, "Sayım Adet", '=COUNTIFS(tblAkis[Durum],"SAYIM",tblAkis[StokAdet],">0")', CATI),
        (9, "ABC Adet", '=COUNTIFS(tblAkis[Durum],"ABC",tblAkis[StokAdet],">0")', CATI),
        (10, "Ölü Stok Adet", "=soa_oluAdet", CATI),
        (11, "Stok Günü Sapma", "=soa_stokGunuSapma", GUN),
        (12, "Yetim Kayıt", "=soa_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=soa_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=soa_zincirKirik", CATI),
        (15, "Tornado Zirve", "=soa_tornadoZirve", TL),
        (16, "Senaryo Farkı", "=soa_senaryoKarsilastirma", TL),
        (17, "Tahmin Aralık", "=soa_tahminAralik", CATI),
        (18, "HHI Yoğunlaşma", "=soa_hhi", YÜZDE),
        (19, "Kalite Skoru", "=soa_kaliteSkor", CATI),
        (20, "P90 Yüzdelik", "=soa_yuzdelikP90", CATI),
        (21, "Kapalı Düşüm", "=soa_kapaliDusum", CATI),
        (22, "Senaryo Motor", "=soa_senaryoMotor", None),
        (23, "Karar Metni", "=soa_kararMetni", None),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=soa_yorumOlu"),
        (8, "=soa_yorumSiparis"),
        (9, "=soa_yorumSapma"),
        (10, "=soa_yorumTornado"),
        (11, "=soa_yorumSenaryo"),
        (12, "=soa_yorumHhi"),
        (13, "=soa_yorumKalite"),
        (14, "=soa_yorumSenMotor"),
        (15, "=soa_yorumTahmin"),
        (16, "=soa_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Sayım", 7), ("Durum ABC", 8), ("Durum Aksiyon", 7),
        ("Durum Kapalı", 6),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Yaş Kova")
    h(ws, 26, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 5), ("8-30", 8), ("31-90", 7), ("90+", 8),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Sipariş")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 40 + i * 12, sayi=CATI)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Nakit")
    for i, (ad, t) in enumerate([
        ("Temkinli", 420000), ("Baz", 510000), ("İyimser", 590000),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=30), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=30))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Stok Günü Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık Sipariş"
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
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Ölü", 8),
    ], 27):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 26, 16, "Ay")
    h(ws, 26, 17, "Ölü Nakit")
    for i in range(6):
        h(ws, 27 + i, 16, i + 1, sayi=CATI)
        h(ws, 27 + i, 17, 35000 + i * 8000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = LineChart()
    c6.title = "Aylık Ölü Stok Nakit"
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
    c8.add_data(Reference(ws, min_col=2, min_row=26, max_row=30), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=27, max_row=30))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F83")

    baski_hazirla(ws, "A1:N36", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "SKU kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=soa_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: SKU-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Sinif", "Depo", "MinStok",
                "RiskNotu", "Aktif", "ToplamStok", "OrtStokGunu"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamStok": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[StokAdet]))'
        ),
        "OrtStokGunu": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[StokGunu]),0))'
        ),
    }
    for i, (kid, unvan, sinif, depo) in enumerate(SKU_KARTLARI):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="SKU-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="SKU unvanı")
        _sari(ws, r, 3, sinif, baslik="Sınıf", mesaj="Malzeme sınıfı")
        _sari(ws, r, 4, depo, baslik="Depo", mesaj="Depo adı")
        _sari(ws, r, 5, 50 + i * 10, sayi=CATI, baslik="Min Stok", mesaj="Asgari stok adet")
        _sari(ws, r, 6, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=CATI if c == 5 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="SKU-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Sınıf"), (4, "Depo"), (6, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "whole", "0", f"E{ILK}:E{SON}",
              baslik="Min Stok", mesaj="Min stok ≥ 0",
              hata_baslik="Min Stok", hata_mesaj="0 veya üzeri",
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
             "Stok sayım/ABC satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "SkuKod", "SayimTarihi", "SonHareket",
        "OncekiDurum", "Durum", "StokAdet", "BirimMaliyet", "KontrolNakit",
        "OnerilenSiparis", "StokGunu", "AbcSinif", "YorumA", "YorumB",
        "HesaplananNakit", "OluBayrak", "YasGun", "YasKova",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu", "SiparisTl",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[SonHareket]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_soa("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[HesaplananNakit]]",
        "tblAkis[[#This Row],[KontrolNakit]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "HesaplananNakit": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[StokAdet]]*tblAkis[[#This Row],[BirimMaliyet]])'
        ),
        "OluBayrak": (
            'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            'tblAkis[[#This Row],[Durum]]="KAPALI"),"",'
            'IF(tblAkis[[#This Row],[StokGunu]]>soa_esikOluGun,"ÖLÜ",""))'
        ),
        "YasGun": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yas_f})',
        "YasKova": f'IF(tblAkis[[#This Row],[IslemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
        "SiparisTl": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[OnerilenSiparis]]*tblAkis[[#This Row],[BirimMaliyet]])'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="SOA-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["sku"], baslik="SKU Kod", mesaj="Stok kodu")
        _sari(ws, r, 4, s["sayim"], sayi=TARİH, baslik="Sayım Tarihi", mesaj="Sayım tarihi")
        _sari(ws, r, 5, s["son_hareket"], sayi=TARİH, baslik="Son Hareket",
              mesaj="Son hareket tarihi")
        _sari(ws, r, 6, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 7, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 8, s["stok"], sayi=CATI, baslik="Stok Adet", mesaj="Stok adet")
        _sari(ws, r, 9, s["maliyet"], sayi=TL, baslik="Birim Maliyet", mesaj="Birim maliyet TL")
        _sari(ws, r, 10, s["kontrol"], sayi=TL, baslik="Kontrol Nakit", mesaj="Kontrol nakit TL")
        _sari(ws, r, 11, s["onerilen"], sayi=CATI, baslik="Önerilen Sipariş", mesaj="Önerilen sipariş adet")
        _sari(ws, r, 12, s["stok_gunu"], sayi=GUN, baslik="Stok Günü", mesaj="Stokta kalma günü")
        _sari(ws, r, 13, s["abc"], baslik="ABC Sınıf", mesaj="A, B veya C")
        _sari(ws, r, 14, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 15, s["yorum_b"], baslik="Yorum B", mesaj="Serbest takip notu")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 16):
            fmt = TARİH if c in (4, 5) else (
                TL if c in (9, 10) else (
                    GUN if c == 12 else (CATI if c in (8, 11) else None)))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:X{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="SOA-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="SKU Kod", mesaj="Stok kodu",
              hata_baslik="SKU", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Sayım", mesaj="Sayım tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "date", "1", f"E{ILK}:E{SON}",
              baslik="Son Hareket", mesaj="Son hareket tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"G{ILK}:G{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(8, "Stok"), (11, "Sipariş"), (12, "Stok Günü")]:
        dogrulama(ws, "whole", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"I{ILK}:I{SON}",
              baslik="Birim Maliyet", mesaj="≥ 0 TL",
              hata_baslik="Maliyet", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"J{ILK}:J{SON}",
              baslik="Kontrol Nakit", mesaj="≥ 0 TL",
              hata_baslik="Nakit", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeAbc", f"M{ILK}:M{SON}",
              baslik="ABC", mesaj="A, B veya C",
              hata_baslik="Liste", hata_mesaj="A/B/C")
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
    h(ws, 5, 3, "Nakit", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Kuyruk Sayım")
    h(ws, 6, 2,
      '=COUNTIFS(tblAkis[Durum],"SAYIM",tblAkis[StokAdet],">0")', sayi=CATI)
    h(ws, 6, 3,
      '=SUMIFS(tblAkis[HesaplananNakit],tblAkis[Durum],"SAYIM")', sayi=TL)

    h(ws, 7, 1, "Kuyruk ABC")
    h(ws, 7, 2,
      '=COUNTIFS(tblAkis[Durum],"ABC",tblAkis[StokAdet],">0")', sayi=CATI)
    h(ws, 7, 3,
      '=SUMIFS(tblAkis[HesaplananNakit],tblAkis[Durum],"ABC")', sayi=TL)

    h(ws, 8, 1, "Kuyruk Aksiyon")
    h(ws, 8, 2,
      '=COUNTIFS(tblAkis[Durum],"AKSIYON",tblAkis[StokAdet],">0")', sayi=CATI)
    h(ws, 8, 3,
      '=SUMIFS(tblAkis[HesaplananNakit],tblAkis[Durum],"AKSIYON")', sayi=TL)

    h(ws, 9, 1, "Kuyruk Ölü Stok")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[OluBayrak],"ÖLÜ",tblAkis[Durum],"<>KAPALI")', sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[HesaplananNakit],tblAkis[OluBayrak],"ÖLÜ",'
      'tblAkis[Durum],"<>KAPALI")', sayi=TL)

    h(ws, 12, 1, "Kapalı — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 1, "Kapalı")
    h(ws, 13, 2, '=COUNTIF(tblAkis[Durum],"KAPALI")', sayi=CATI)
    h(ws, 13, 3, '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[HesaplananNakit])', sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=soa_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-90", "90+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[HesaplananNakit])', sayi=TL)

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
        ("H03", "Ölü stok nakit",
         '=SUMIFS(tblAkis[HesaplananNakit],tblAkis[OluBayrak],"ÖLÜ",'
         'tblAkis[Durum],"<>KAPALI")'),
        ("H04", "Stok adet toplam", "=SUM(tblAkis[StokAdet])"),
        ("H05", "Kontrol nakit toplam", "=SUM(tblAkis[KontrolNakit])"),
        ("H06", "Hesaplanan nakit toplam", "=SUM(tblAkis[HesaplananNakit])"),
        ("H07", "Sayım adet", '=COUNTIF(tblAkis[Durum],"SAYIM")'),
        ("H08", "ABC adet", '=COUNTIF(tblAkis[Durum],"ABC")'),
        ("H09", "Aksiyon adet", '=COUNTIF(tblAkis[Durum],"AKSIYON")'),
        ("H10", "Ölü bayrak adet",
         '=COUNTIFS(tblAkis[OluBayrak],"ÖLÜ",tblAkis[Durum],"<>KAPALI")'),
        ("H11", "Kapalı adet", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        ("H12", "Açık kuyruk adet", "=C12+C13+C14"),
        ("H13", "Kapalı düşüm adet", "=C16"),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H19", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H20", "Yaş 31-90", '=COUNTIF(tblAkis[YasKova],"31-90")'),
        ("H21", "Yaş 90+", '=COUNTIF(tblAkis[YasKova],"90+")'),
        ("H22", "Ort stok günü",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[StokGunu]),0)"),
        ("H23", "Ölü stok adet", "=C15"),
        ("H24", "Ölü stok nakit (T1)", "=C8"),
        ("H25", "Önerilen sipariş ort",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[OnerilenSiparis]),0)"),
        ("H26", "Stok günü sapma", "=C27-soa_hedefStokGunu"),
        ("H27", "Toplam stok nakit", "=C11"),
        ("H28", "Temkinli senaryo", "=C32*soa_senaryoTemkinli"),
        ("H29", "Baz senaryo", "=C32*soa_senaryoBaz"),
        ("H30", "İyimser senaryo", "=C32*soa_senaryoIyimser"),
        ("H31", "Senaryo farkı", "=C35-C33"),
        ("H32", "Tornado (oran×nakit)", "=C32*soa_tornadoOran"),
        ("H33", "HHI proxy",
         "=IFERROR((MAX(C12,C13,C14)/MAX(C17,1))^2,0)"),
        ("H34", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,100-C19*10-C20*10-C21*10-C28*0.1))"),
        ("H35", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[StokAdet],soa_yuzdelikOran),0)"),
        ("H36", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[StokAdet],"
         "tblAkis[KayitDolu]),C40)"),
        ("H37", "Tahmin aralık", "=IFERROR((C40+C41)/2,0)"),
        ("H38", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[OnerilenSiparis],soa_yuzdelikOran),0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Min stok toplam", "=SUM(tblKartlar[MinStok])"),
        ("H41", "Sipariş TL toplam",
         "=IFERROR(SUMIF(tblAkis[KayitDolu],1,tblAkis[SiparisTl]),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Sayım",C12,"ABC",C13,"Aksiyon",C14,'
         '"Ölü",C15)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C23,C24,C25,C26)'),
        ("H44", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C33,"Baz",C34,"İyimser",C35)'),
        ("H45", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H46", "Ölü kritik bayrak",
         '=IF(C8>soa_esikOluNakit,1,0)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "nakit", "senaryo", "tornado", "fark", "tl")) and "oran" not in acik.lower() and "adet" not in acik.lower() and "metin" not in acik.lower() and "özet" not in acik.lower() and "bayrak" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "hhi" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "gün" in acik.lower() or "sapma" in acik.lower():
            fmt = GUN
        if "kalite" in acik.lower() or "bayrak" in acik.lower():
            fmt = CATI
        if "hhi" in acik.lower() or "hazırlık" in acik.lower():
            fmt = YÜZDE
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 55, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 56, 1, "Temkinli")
    h(ws, 56, 2, "=soa_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 56, 3, "=C33", sayi=TL)
    h(ws, 57, 1, "Baz")
    h(ws, 57, 2, "=soa_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 57, 3, "=C34", sayi=TL)
    h(ws, 58, 1, "İyimser")
    h(ws, 58, 2, "=soa_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 58, 3, "=C35", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Stok Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Ölü Stok Nakit", "=soa_oluStokNakit", TL),
        (10, "Önerilen Sipariş", "=soa_onerilenSiparis", CATI),
        (11, "Stok Günü Sapma", "=soa_stokGunuSapma", GUN),
        (12, "Yetim", "=soa_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=soa_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=soa_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=soa_tahminAralik", CATI),
        (16, "Kanıt Özeti", "=soa_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Durum haritası AYARLAR'dadır (SOA-001). Geçersiz geçiş engellenmez, görünür kılınır "
      "(SOA-002). Stok adet×maliyet ↔ kontrol nakit zinciri toleransla denetlenir (SOA-003). "
      "Kapalı kayıtlar kuyruktan düşer (SOA-004). Ölü stok eşikleri kararı belirler (SOA-005). "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Yetim SKU kayıtlarını KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 23, 1, "2. Ölü stok gün eşiğini aşan kalemlerde eritme planı açın.")
    h(ws, 24, 1, "3. Zincir kırık satırlarda adet×maliyet ↔ kontrol nakit farkını açıklayın.")
    h(ws, 26, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H26", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Sayım", '=COUNTIF(tblAkis[Durum],"SAYIM")'),
        (8, "ABC", '=COUNTIF(tblAkis[Durum],"ABC")'),
        (9, "Aksiyon", '=COUNTIF(tblAkis[Durum],"AKSIYON")'),
        (10, "Kapalı", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        (11, "Ölü bayrak", '=COUNTIFS(tblAkis[OluBayrak],"ÖLÜ")'),
        (12, "Açık toplam", "=B7+B8+B9"),
        (13, "Yetim", "=soa_yetimKayit"),
        (14, "Geçersiz", "=soa_gecersizGecis"),
        (15, "Zincir kırık", "=soa_zincirKirik"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=soa_kaliteSkor"),
        (18, "Ölü stok nakit", "=soa_oluStokNakit"),
        (19, "Önerilen sipariş", "=soa_onerilenSiparis"),
        (20, "Kapalı düşüm", "=soa_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r == 18 else (CATI if r != 16 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} SKU kartı ve {DEMO_AKIS} stok satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Sayım/ABC/Aksiyon/Kapalı — kapalılar kuyruktan düşer")
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

    h(ws, 3, 9, "ABC Sınıf", kalin=True)
    h(ws, 6, 9, "Abc", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["A", "B", "C"], 7):
        _sari(ws, i, 9, b, baslik="ABC", mesaj="ABC sınıfı")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası + SOA kuralları", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 100, "TL", "İş kuralı — adet×maliyet/kontrol nakit sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("soa_esikOluGun", 90, "gün", "Ölü stok politikası SOA-005", "01.01.2026",
         "Ölü stok gün eşiği"),
        ("soa_esikOluNakit", 150000, "TL", "Ölü stok politikası", "01.01.2026",
         "ÖLÜ STOK ERİT nakit eşiği"),
        ("soa_esikOluAdet", 5, "adet", "Ölü stok politikası", "01.01.2026",
         "ÖLÜ STOK ERİT adet eşiği"),
        ("soa_esikSiparis", 25, "adet", "Sipariş politikası", "01.01.2026",
         "SİPARİŞ ort. eşiği"),
        ("soa_kovaEsik1", 7, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("soa_kovaEsik2", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("soa_kovaEsik3", 90, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("soa_hedefStokGunu", 45, "gün", "Stok günü modeli", "01.01.2026", "Hedef stok günü"),
        ("soa_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("soa_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("soa_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("soa_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("soa_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("soa_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("soa_azamiMiktar", 500000, "adet", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("soa_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi SOA-001",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("soa_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("soa_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06 SOA-004",
         "01.01.2026", "Kapalı düşüm kuralı"),
        ("soa_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("soa_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("soa_firmaUnvan", "Örnek Stok A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("soa_kuralSOA001", "SOA-001", "kural", "Durum haritası zorunlu", "01.01.2026",
         "Durum makinesi AYARLAR'da"),
        ("soa_kuralSOA002", "SOA-002", "kural", "Geçersiz geçiş görünür", "01.01.2026",
         "Engellemez, bayrak üretir"),
        ("soa_kuralSOA003", "SOA-003", "kural", "Adet×maliyet=kontrol nakit zinciri", "01.01.2026",
         "ZİNCİR KIRIK"),
        ("soa_kuralSOA004", "SOA-004", "kural", "Kapalı kuyruktan düşer", "01.01.2026",
         "A06 kapalı düşüm"),
        ("soa_kuralSOA005", "SOA-005", "kural", "Ölü stok eşik kararları", "01.01.2026",
         "ÖLÜ STOK ERİT / SİPARİŞ"),
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
    tablo_ekle(ws, "tblDurumHaritasi", "H38:L42", dh_bas)

    h(ws, 46, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 47, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 47, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (48, "soa_oluStokNakit", "=HESAP!C8"),
        (49, "soa_onerilenSiparis", "=HESAP!C30"),
        (50, "soa_oluAdet", "=HESAP!C28"),
        (51, "soa_stokGunuSapma", "=HESAP!C31"),
        (52, "soa_yetimKayit", "=HESAP!C19"),
        (53, "soa_gecersizGecis", "=HESAP!C20"),
        (54, "soa_zincirKirik", "=HESAP!C21"),
        (55, "soa_tahminAralik", "=HESAP!C42"),
        (56, "soa_senaryoKarsilastirma", "=HESAP!C36"),
        (57, "soa_tornadoZirve", "=HESAP!C37"),
        (58, "soa_hhi", "=HESAP!C38"),
        (59, "soa_kaliteSkor", "=HESAP!C39"),
        (60, "soa_yuzdelikP90", "=HESAP!C43"),
        (61, "soa_senaryoMotor", "=HESAP!C49"),
        (62, "soa_kapaliDusum", "=HESAP!C18"),
        (63, "soa_kuyrukOzeti", "=HESAP!C47"),
        (64, "soa_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"ÖlüNakit",TEXT(HESAP!C8,"0"),'
         '"SiparişOrt",HESAP!C30,"Yetim",HESAP!C19,"Zincir",HESAP!C21,"ÖlüAdet",HESAP!C28)'),
        (65, "soa_sonrakiKimlik",
         sonraki_kimlik_formul("SKU", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
        (66, "soa_kararMetni",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Karar",PANO!B4,"— ölü nakit",TEXT(soa_oluStokNakit,"0"),'
         '"sipariş ort",soa_onerilenSiparis)'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 68, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (69, "soa_yorumOlu",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ölü stok nakit",TEXT(soa_oluStokNakit,"0"),"TL")'),
        (70, "soa_yorumSiparis",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Önerilen sipariş ort",soa_onerilenSiparis,"adet")'),
        (71, "soa_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Stok günü sapma",soa_stokGunuSapma,"gün")'),
        (72, "soa_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(soa_tornadoZirve,"0.00"),"TL")'),
        (73, "soa_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(soa_senaryoKarsilastirma,"0.00"),"TL")'),
        (74, "soa_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(soa_hhi,"0.0%"))'),
        (75, "soa_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",soa_kaliteSkor)'),
        (76, "soa_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",soa_senaryoMotor)'),
        (77, "soa_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",soa_tahminAralik)'),
        (78, "soa_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 sipariş",soa_yuzdelikP90)'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına SKU kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her stok sayım/ABC satırını yazın; durum kolonunu listeden seçin.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — uyarı üretir (makrosuz kural, SOA-002).",
        "5. Kapalı durum kuyruklardan düşer (SOA-004).",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma son harekete göredir.",
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
    for r in range(6, 34):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 11):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Durum", mesaj="Durum kimliği",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    for r in range(7, 12):
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
        "B4", CellIsRule(operator="equal", formula=['"BEKLE"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"SİPARİŞ"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"ÖLÜ STOK ERİT"'], fill=kirmizi, font=kf))
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
        ("Q", '"ÖLÜ"', kirmizi, kf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("ABC", sari), ("AKSIYON", kirmizi), ("SAYIM", yesil)]:
        ws.conditional_formatting.add(
            f"G{ILK}:G{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["KUYRUKLAR"]
    for r in range(6, 10):
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
        "B7", CellIsRule(operator="equal", formula=['"BEKLE"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"SİPARİŞ"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"ÖLÜ STOK ERİT"'], fill=kirmizi, font=kf))
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
        CellIsRule(operator="greaterThan", formula=["100"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans",
        "soa_esikOluGun", "soa_esikOluNakit", "soa_esikOluAdet", "soa_esikSiparis",
        "soa_kovaEsik1", "soa_kovaEsik2", "soa_kovaEsik3",
        "soa_hedefStokGunu",
        "soa_senaryoTemkinli", "soa_senaryoBaz", "soa_senaryoIyimser",
        "soa_tornadoOran", "soa_yuzdelikOran", "soa_olcekHedef", "soa_azamiMiktar",
        "soa_durumHaritasi", "soa_kuyrukOzeti_metin", "soa_kapaliDusum_metin",
        "soa_kanitRaporu_metin", "soa_dosyaSurumu", "soa_firmaUnvan",
        "soa_kuralSOA001", "soa_kuralSOA002", "soa_kuralSOA003",
        "soa_kuralSOA004", "soa_kuralSOA005",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "soa_oluStokNakit": 48, "soa_onerilenSiparis": 49, "soa_oluAdet": 50,
        "soa_stokGunuSapma": 51, "soa_yetimKayit": 52, "soa_gecersizGecis": 53,
        "soa_zincirKirik": 54, "soa_tahminAralik": 55,
        "soa_senaryoKarsilastirma": 56, "soa_tornadoZirve": 57, "soa_hhi": 58,
        "soa_kaliteSkor": 59, "soa_yuzdelikP90": 60, "soa_senaryoMotor": 61,
        "soa_kapaliDusum": 62, "soa_kuyrukOzeti": 63, "soa_kanitRaporu": 64,
        "soa_sonrakiKimlik": 65, "soa_kararMetni": 66,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "soa_yorumOlu": 69, "soa_yorumSiparis": 70, "soa_yorumSapma": 71,
        "soa_yorumTornado": 72, "soa_yorumSenaryo": 73, "soa_yorumHhi": 74,
        "soa_yorumKalite": 75, "soa_yorumSenMotor": 76, "soa_yorumTahmin": 77,
        "soa_yorumP90": 78,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$G$7:$G$8")
    ad_ekle(wb, "ListeAbc", "LISTELER!$I$7:$I$9")


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
    dosya_ad = "StokOptimizasyonAbcOluStokNakit.xlsx"
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
