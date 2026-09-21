#!/usr/bin/env python3
"""Mini-MRP (BOM + Malzeme İhtiyacı + Kapasite) — A3 üretim betiği (manda v6)."""

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

URUN_AD = "Mini-MRP (BOM + Malzeme İhtiyacı + Kapasite)"
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
    {"durum_id": "PLAN", "durum_adi": "Plan", "sira": 1,
     "izinli_sonraki_durumlar": "TEDARIK", "kategori": "acik"},
    {"durum_id": "TEDARIK", "durum_adi": "Tedarik", "sira": 2,
     "izinli_sonraki_durumlar": "URETIM|PLAN", "kategori": "acik"},
    {"durum_id": "URETIM", "durum_adi": "Üretim", "sira": 3,
     "izinli_sonraki_durumlar": "TAMAM|TEDARIK", "kategori": "acik"},
    {"durum_id": "TAMAM", "durum_adi": "Tamam", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "PLAN|TEDARIK",
    "TEDARIK|URETIM",
    "TEDARIK|PLAN",
    "URETIM|TAMAM",
    "URETIM|TEDARIK",
]

MAMULLER = [
    ("MAM-00001", "Hidrolik Pompa HP-100", "Hat-A", 40),
    ("MAM-00002", "Dişli Kutusu DK-50", "Hat-B", 36),
    ("MAM-00003", "Valf Bloğu VB-20", "Hat-A", 48),
    ("MAM-00004", "Rulman Ünitesi RU-10", "Hat-C", 32),
    ("MAM-00005", "Motor Flanşı MF-30", "Hat-B", 44),
    ("MAM-00006", "Conta Seti CS-05", "Hat-C", 28),
    ("MAM-00007", "Şaft Komple SK-15", "Hat-A", 40),
    ("MAM-00008", "Kaplin CK-08", "Hat-B", 36),
]

BILESENLER = [
    "CEL-001", "PLK-002", "RUL-003", "CON-004",
    "VID-005", "YAG-006", "FLN-007", "KAB-008",
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Net ihtiyaç ve kapasite kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["PLAN", "TEDARIK", "URETIM", "TAMAM",
                "TEDARIK", "URETIM", "PLAN", "TEDARIK"]
    oncekiler = ["", "PLAN", "TEDARIK", "URETIM",
                 "PLAN", "TEDARIK", "PLAN", "TEDARIK"]
    for i in range(1, DEMO_AKIS + 1):
        kart = MAMULLER[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "MAM-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: PLAN → URETIM
            o, d = "PLAN", "URETIM"
        if i == 20:  # geçersiz: URETIM → PLAN
            o, d = "URETIM", "PLAN"
        siparis = 10 + (i % 8) * 5
        birim = 1 + (i % 4)
        stok = max(0, siparis * birim - (i * 3) % 20)
        brut_hesap = siparis * birim
        brut_beklenen = brut_hesap
        if i in (12, 19):  # zincir kırık
            brut_beklenen = brut_hesap + 15
        tedarik = 3 + (i % 10)
        is_emri = 2 + (i % 6) * 0.5
        plan_tar = RAPOR_TARIHI - timedelta(days=(i * 2) % 40)
        satirlar.append({
            "id": f"IE-{i:05d}",
            "kart": kart,
            "bilesen": BILESENLER[(i - 1) % len(BILESENLER)],
            "birim": birim,
            "stok": stok,
            "siparis": siparis,
            "tedarik": tedarik,
            "is_emri": is_emri,
            "onceki": o,
            "durum": d,
            "brut_beklenen": brut_beklenen,
            "plan_tar": plan_tar,
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"İş emri notu {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Mamul BOM satırlarını stok, sipariş ve kapasite ile patlatır; net ihtiyaç ve "
      "kapasite yükünü görünür kılar; ÜRET / SATIN AL / KAPASİTE AŞIMI kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Plan → Tedarik → Üretim → Tamam",
        "Brüt/net ihtiyaç ve önerilen sipariş motoru",
        "Kapasite yükü % ve aşım bayrağı",
        "Geçersiz geçiş ve yetim mamul uyarıları",
        "Dönemsel malzeme kanıt raporu — yönetici dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Üretim planlama — net ihtiyaç ve iş emri kuyruğu",
        "Satın alma / fabrika müdürü — ÜRET / SATIN AL kararı",
        "SMMM / mali işler — malzeme kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a mamul kartı girin. 2) AKIS'e BOM/iş emri satırlarını yazın. "
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
      'IF(mmr_kapasiteAsim>0,"KAPASİTE AŞIMI",'
      'IF(mmr_satinAlAdet>mmr_esikSatinAl,"SATIN AL","ÜRET")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: ÜRET | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Plan Kuyruk Adet", '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[KayitDolu],1)', CATI),
        (7, "Tedarik Adet", '=COUNTIFS(tblAkis[Durum],"TEDARIK",tblAkis[KayitDolu],1)', CATI),
        (8, "Üretim Adet", '=COUNTIFS(tblAkis[Durum],"URETIM",tblAkis[KayitDolu],1)', CATI),
        (9, "Net İhtiyaç", "=mmr_netIhtiyac", CATI),
        (10, "Önerilen Sipariş", "=mmr_onerilenSiparis", CATI),
        (11, "Kapasite Yükü", "=mmr_kapasiteYuku", YÜZDE),
        (12, "Yetim Kayıt", "=mmr_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=mmr_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=mmr_zincirKirik", CATI),
        (15, "Kapasite Aşım", "=mmr_kapasiteAsim", CATI),
        (16, "Satın Al Adet", "=mmr_satinAlAdet", CATI),
        (17, "Tahmin Aralık", "=mmr_tahminAralik", CATI),
        (18, "Senaryo Farkı", "=mmr_senaryoKarsilastirma", CATI),
        (19, "Tornado Zirve", "=mmr_tornadoZirve", CATI),
        (20, "HHI Yoğunlaşma", "=mmr_hhi", YÜZDE),
        (21, "Kalite Skoru", "=mmr_kaliteSkor", CATI),
        (22, "P90 Yüzdelik", "=mmr_yuzdelikP90", CATI),
        (23, "Kapalı Düşüm", "=mmr_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=mmr_yorumNet"),
        (8, "=mmr_yorumKapasite"),
        (9, "=mmr_yorumSapma"),
        (10, "=mmr_yorumTornado"),
        (11, "=mmr_yorumSenaryo"),
        (12, "=mmr_yorumHhi"),
        (13, "=mmr_yorumKalite"),
        (14, "=mmr_yorumSenMotor"),
        (15, "=mmr_yorumTahmin"),
        (16, "=mmr_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Plan", 7), ("Durum Tedarik", 8), ("Durum Üretim", 7),
        ("Durum Tamam", 6),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Bileşen")
    h(ws, 26, 5, "Net")
    for i, (ad, adet) in enumerate([
        ("CEL-001", 12), ("PLK-002", 18), ("RUL-003", 9), ("CON-004", 15),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Netİhtiyaç")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 40 + i * 8, sayi=CATI)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "İhtiyaç")
    for i, (ad, t) in enumerate([
        ("Temkinli", 180), ("Baz", 220), ("İyimser", 260),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=CATI)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=30), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=30))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Bileşen Net İhtiyaç"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık Net İhtiyaç"
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
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Aşım", 3),
    ], 27):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 26, 16, "Ay")
    h(ws, 26, 17, "Yük")
    for i in range(6):
        h(ws, 27 + i, 16, i + 1, sayi=CATI)
        h(ws, 27 + i, 17, 0.55 + i * 0.07, sayi=YÜZDE)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = LineChart()
    c6.title = "Aylık Kapasite Yükü"
    c6.add_data(Reference(ws, min_col=17, min_row=26, max_row=32), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=27, max_row=32))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F68")

    c7 = PieChart()
    c7.title = "Bileşen Payı"
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
    alt_bant(ws, 3, "Mamul kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=mmr_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: MAM-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "MamulAd", "HatAdi", "KapasiteSaat",
                "MinStok", "Aktif", "ToplamSiparis", "OrtYuk"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamSiparis": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[SiparisMiktari]))'
        ),
        "OrtYuk": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[KapasiteYukuPct]),0))'
        ),
    }
    for i, (kid, ad, hat, kap) in enumerate(MAMULLER):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="MAM-##### formatında")
        _sari(ws, r, 2, ad, baslik="Mamul Adı", mesaj="Mamul tanımı")
        _sari(ws, r, 3, hat, baslik="Hat", mesaj="Üretim hattı")
        _sari(ws, r, 4, kap, sayi=CATI, baslik="Kapasite Saat", mesaj="Haftalık kapasite saat")
        _sari(ws, r, 5, 5 + i, sayi=CATI, baslik="Min Stok", mesaj="Asgari stok")
        _sari(ws, r, 6, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 7):
            _sari(ws, r, c, None, sayi=CATI if c in (4, 5) else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:H{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="MAM-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Mamul"), (3, "Hat")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(4, "Kapasite"), (5, "Min Stok")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"F{ILK}:F{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["B"].width = 24


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=24)
    alt_bant(ws, 3,
             "BOM / iş emri satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "BilesenKod", "BirimIhtiyac", "StokMiktar",
        "SiparisMiktari", "TedarikSuresiGun", "IsEmriSuresiSaat", "OncekiDurum", "Durum",
        "BrutBeklenen", "PlanTarihi", "YorumA", "YorumB",
        "BrutIhtiyac", "NetIhtiyac", "OnerilenSiparis", "KapasiteYukuPct",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu", "YasGun",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[PlanTarihi]]", "raporTarihi").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[BrutIhtiyac]]",
        "tblAkis[[#This Row],[BrutBeklenen]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "BrutIhtiyac": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'tblAkis[[#This Row],[SiparisMiktari]]*tblAkis[[#This Row],[BirimIhtiyac]])'
        ),
        "NetIhtiyac": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'MAX(0,tblAkis[[#This Row],[BrutIhtiyac]]-tblAkis[[#This Row],[StokMiktar]]))'
        ),
        "OnerilenSiparis": (
            'IF(tblAkis[[#This Row],[IslemId]]="","",'
            'IF(OR(tblAkis[[#This Row],[Durum]]="TAMAM"),0,'
            'tblAkis[[#This Row],[NetIhtiyac]]))'
        ),
        "KapasiteYukuPct": (
            'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            'tblAkis[[#This Row],[Durum]]="TAMAM"),"",'
            'IFERROR(tblAkis[[#This Row],[IsEmriSuresiSaat]]/'
            'MAX(1,SUMIF(tblKartlar[KartId],tblAkis[[#This Row],[KaynakKartId]],'
            'tblKartlar[KapasiteSaat])),0))'
        ),
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
        "YasGun": (
            f'IF(tblAkis[[#This Row],[IslemId]]="","",'
            f'IF(tblAkis[[#This Row],[Durum]]="TAMAM",0,{yas_f}))'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="IE-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["bilesen"], baslik="Bileşen Kod", mesaj="BOM bileşen kodu")
        _sari(ws, r, 4, s["birim"], sayi=CATI, baslik="Birim İhtiyaç", mesaj="Mamul başına bileşen")
        _sari(ws, r, 5, s["stok"], sayi=CATI, baslik="Stok", mesaj="Mevcut stok miktarı")
        _sari(ws, r, 6, s["siparis"], sayi=CATI, baslik="Sipariş Miktarı", mesaj="Mamul sipariş adedi")
        _sari(ws, r, 7, s["tedarik"], sayi=GUN, baslik="Tedarik Süresi", mesaj="Tedarik süresi gün")
        _sari(ws, r, 8, s["is_emri"], sayi=CATI, baslik="İş Emri Süresi", mesaj="İş emri süresi saat")
        _sari(ws, r, 9, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 10, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 11, s["brut_beklenen"], sayi=CATI, baslik="Brüt Beklenen",
              mesaj="Beklenen brüt ihtiyaç (zincir)")
        _sari(ws, r, 12, s["plan_tar"], sayi=TARİH, baslik="Plan Tarihi", mesaj="Plan tarihi")
        _sari(ws, r, 13, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 14, s["yorum_b"], baslik="Yorum B", mesaj="Serbest takip notu")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 15):
            fmt = TARİH if c == 12 else (
                GUN if c == 7 else (CATI if c in (4, 5, 6, 8, 11) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:W{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="IE-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Bileşen", mesaj="Bileşen kodu",
              hata_baslik="Bileşen", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    for c, ad in [(4, "Birim"), (5, "Stok"), (6, "Sipariş"), (7, "Tedarik"),
                  (8, "İş Emri"), (11, "Brüt Beklenen")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeDurumlar", f"I{ILK}:I{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"J{ILK}:J{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "date", "1", f"L{ILK}:L{SON}",
              baslik="Plan", mesaj="Plan tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeYorumAB", f"M{ILK}:M{SON}",
              baslik="Yorum A", mesaj="A veya B",
              hata_baslik="Liste", hata_mesaj="A veya B")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 24)})
    ws.column_dimensions["N"].width = 16
    ws.column_dimensions["T"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Net İhtiyaç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Kuyruk Plan")
    h(ws, 6, 2, '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[KayitDolu],1)', sayi=CATI)
    h(ws, 6, 3, '=SUMIF(tblAkis[Durum],"PLAN",tblAkis[NetIhtiyac])', sayi=CATI)

    h(ws, 7, 1, "Kuyruk Tedarik")
    h(ws, 7, 2, '=COUNTIFS(tblAkis[Durum],"TEDARIK",tblAkis[KayitDolu],1)', sayi=CATI)
    h(ws, 7, 3, '=SUMIF(tblAkis[Durum],"TEDARIK",tblAkis[NetIhtiyac])', sayi=CATI)

    h(ws, 8, 1, "Kuyruk Üretim")
    h(ws, 8, 2, '=COUNTIFS(tblAkis[Durum],"URETIM",tblAkis[KayitDolu],1)', sayi=CATI)
    h(ws, 8, 3, '=SUMIF(tblAkis[Durum],"URETIM",tblAkis[NetIhtiyac])', sayi=CATI)

    h(ws, 9, 1, "Satın Al (net>0 ve tedarik≥eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[NetIhtiyac],">0",tblAkis[TedarikSuresiGun],">="&mmr_esikTedarikGun,'
      'tblAkis[Durum],"<>TAMAM")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[NetIhtiyac],tblAkis[NetIhtiyac],">0",'
      'tblAkis[TedarikSuresiGun],">="&mmr_esikTedarikGun,tblAkis[Durum],"<>TAMAM")',
      sayi=CATI)

    h(ws, 11, 1, "Kapalı (Tamam) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Tamam")
    h(ws, 12, 2, '=COUNTIF(tblAkis[Durum],"TAMAM")', sayi=CATI)
    h(ws, 12, 3, '=SUMIF(tblAkis[Durum],"TAMAM",tblAkis[NetIhtiyac])', sayi=CATI)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=mmr_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Kapasite Yük Bandı", kalin=True, yazi=KOYU_LACIVERT)
    for i, (ad, form) in enumerate([
        ("≤%70", '=COUNTIFS(tblAkis[KapasiteYukuPct],"<="&0.7,tblAkis[KayitDolu],1)'),
        ("%70-100", '=COUNTIFS(tblAkis[KapasiteYukuPct],">0.7",tblAkis[KapasiteYukuPct],"<=1",'
                    'tblAkis[KayitDolu],1)'),
        (">%100", '=COUNTIFS(tblAkis[KapasiteYukuPct],">1",tblAkis[KayitDolu],1)'),
        ("Boş/Tamam", '=COUNTBLANK(tblAkis[KapasiteYukuPct])'),
    ], 18):
        h(ws, i, 1, ad)
        h(ws, i, 2, form, sayi=CATI)

    genislik(ws, {"A": 40, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam işlem", '=COUNTA(tblAkis[IslemId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Brüt ihtiyaç toplam", "=SUM(tblAkis[BrutIhtiyac])"),
        ("H04", "Stok toplam", "=SUM(tblAkis[StokMiktar])"),
        ("H05", "Sipariş toplam", "=SUM(tblAkis[SiparisMiktari])"),
        ("H06", "Plan adet", '=COUNTIF(tblAkis[Durum],"PLAN")'),
        ("H07", "Tedarik adet", '=COUNTIF(tblAkis[Durum],"TEDARIK")'),
        ("H08", "Üretim adet", '=COUNTIF(tblAkis[Durum],"URETIM")'),
        ("H09", "Tamam adet", '=COUNTIF(tblAkis[Durum],"TAMAM")'),
        ("H10", "Açık kuyruk adet", "=C11+C12+C13"),
        ("H11", "Kapalı düşüm adet", "=C14"),
        ("H12", "Net ihtiyaç toplam", "=SUM(tblAkis[NetIhtiyac])"),
        ("H13", "Önerilen sipariş", "=SUM(tblAkis[OnerilenSiparis])"),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Kapasite aşım adet",
         '=COUNTIFS(tblAkis[KapasiteYukuPct],">"&mmr_esikKapasite,tblAkis[KayitDolu],1)'),
        ("H19", "Ort kapasite yükü",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[KapasiteYukuPct]),0)"),
        ("H20", "Satın al adet",
         '=COUNTIFS(tblAkis[NetIhtiyac],">0",tblAkis[TedarikSuresiGun],">="&mmr_esikTedarikGun,'
         'tblAkis[Durum],"<>TAMAM")'),
        ("H21", "Ort tedarik gün",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[TedarikSuresiGun]),0)"),
        ("H22", "Ort iş emri saat",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[IsEmriSuresiSaat]),0)"),
        ("H23", "İhtiyaç sapma", "=C8-C9"),
        ("H24", "Brüt-beklenen fark",
         "=SUM(tblAkis[BrutIhtiyac])-SUM(tblAkis[BrutBeklenen])"),
        ("H25", "Temkinli senaryo", "=C17*mmr_senaryoTemkinli"),
        ("H26", "Baz senaryo", "=C17*mmr_senaryoBaz"),
        ("H27", "İyimser senaryo", "=C17*mmr_senaryoIyimser"),
        ("H28", "Senaryo farkı", "=C32-C30"),
        ("H29", "Tornado (oran×net)", "=C17*mmr_tornadoOran"),
        ("H30", "HHI proxy",
         "=IFERROR((MAX(C11,C12,C13)/MAX(C15,1))^2,0)"),
        ("H31", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,100-C19*10-C20*10-C21*10-C23*5))"),
        ("H32", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[NetIhtiyac],mmr_yuzdelikOran),0)"),
        ("H33", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[NetIhtiyac],"
         "tblAkis[KayitDolu]),C37)"),
        ("H34", "Tahmin aralık", "=IFERROR((C37+C38)/2,0)"),
        ("H35", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[BrutIhtiyac],mmr_yuzdelikOran),0)"),
        ("H36", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H37", "Kapasite saat toplam", "=SUM(tblKartlar[KapasiteSaat])"),
        ("H38", "Kapasite kullanım", "=IFERROR(C17/MAX(C42,1),0)"),
        ("H39", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Plan",C11,"Tedarik",C12,"Üretim",C13,"SatınAl",C25)'),
        ("H40", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C30,"Baz",C31,"İyimser",C32)'),
        ("H41", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H42", "MMR-001 net≥0", '=IF(MIN(tblAkis[NetIhtiyac])>=0,"UYGUN","İHLAL")'),
        ("H43", "MMR-002 kapasite bayrak",
         '=IF(C23>0,"AŞIM","UYGUN")'),
        ("H44", "MMR-003 zincir",
         '=IF(C21=0,"UYGUN","KIRIK")'),
        ("H45", "MMR-004 yetim",
         '=IF(C19=0,"UYGUN","YETİM")'),
        ("H46", "MMR-005 karar yolu",
         '=IF(C6=0,"VERİ YOK",IF(C23>0,"KAPASİTE AŞIMI",'
         'IF(C25>mmr_esikSatinAl,"SATIN AL","ÜRET")))'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = YÜZDE if any(x in acik.lower() for x in (
            "yük", "oran", "hhi", "hazırlık", "kullanım")) else (
            None if any(x in acik.lower() for x in (
                "metin", "özet", "mmr-", "bayrak", "yol", "zincir", "yetim"))
            and "adet" not in acik.lower() else CATI)
        if "gün" in acik.lower() or ("saat" in acik.lower() and "toplam" not in acik.lower()):
            fmt = GUN if "gün" in acik.lower() else CATI
        if "kalite" in acik.lower():
            fmt = CATI
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 55, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 56, 1, "Temkinli")
    h(ws, 56, 2, "=mmr_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 56, 3, "=C30", sayi=CATI)
    h(ws, 57, 1, "Baz")
    h(ws, 57, 2, "=mmr_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 57, 3, "=C31", sayi=CATI)
    h(ws, 58, 1, "İyimser")
    h(ws, 58, 2, "=mmr_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 58, 3, "=C32", sayi=CATI)

    genislik(ws, {"A": 10, "B": 32, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Malzeme / Kapasite Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Net İhtiyaç", "=mmr_netIhtiyac", CATI),
        (10, "Önerilen Sipariş", "=mmr_onerilenSiparis", CATI),
        (11, "Kapasite Yükü", "=mmr_kapasiteYuku", YÜZDE),
        (12, "Yetim", "=mmr_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=mmr_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=mmr_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=mmr_tahminAralik", CATI),
        (16, "Kanıt Özeti", "=mmr_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Net ihtiyaç = MAX(0, brüt − stok). Kapasite yükü = iş emri süresi / hat kapasitesi. "
      "Kurallar MMR-001..005 (2025-2027). Bu çıktı karar destektir; kesin üretim emri değildir.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 23, 1, "2. Kapasite aşımı olan hatlarda iş emrini kaydırın veya fasona alın.")
    h(ws, 24, 1, "3. Zincir kırık satırlarda brüt ihtiyaç ↔ beklenen farkını açıklayın.")
    h(ws, 26, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H26", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Plan", '=COUNTIF(tblAkis[Durum],"PLAN")'),
        (8, "Tedarik", '=COUNTIF(tblAkis[Durum],"TEDARIK")'),
        (9, "Üretim", '=COUNTIF(tblAkis[Durum],"URETIM")'),
        (10, "Tamam", '=COUNTIF(tblAkis[Durum],"TAMAM")'),
        (11, "Açık toplam", "=B7+B8+B9"),
        (12, "Yetim", "=mmr_yetimKayit"),
        (13, "Geçersiz", "=mmr_gecersizGecis"),
        (14, "Zincir kırık", "=mmr_zincirKirik"),
        (15, "Kapasite aşım", "=mmr_kapasiteAsim"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=mmr_kaliteSkor"),
        (18, "Net ihtiyaç", "=mmr_netIhtiyac"),
        (19, "Önerilen sipariş", "=mmr_onerilenSiparis"),
        (20, "Kapalı düşüm", "=mmr_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True, sayi=CATI if r != 16 else None)
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} mamul kartı ve {DEMO_AKIS} BOM/iş emri satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Plan/Tedarik/Üretim/Tamam — kapalılar kuyruktan düşer")
    h(ws, 8, 1, "Kurallar: MMR-001..005 | Hesap yılları: 2025-2027")
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

    h(ws, 3, 9, "Bileşen", kalin=True)
    h(ws, 6, 9, "Bilesen", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(BILESENLER, 7):
        _sari(ws, i, 9, b, baslik="Bileşen", mesaj="Bileşen listesi")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 5, "adet", "İş kuralı — brüt/beklenen sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("mmr_esikTedarikGun", 7, "gün", "İç politika", "01.01.2026", "Satın al tedarik eşiği"),
        ("mmr_esikSatinAl", 3, "adet", "İç politika", "01.01.2026", "Karar eşiği satın al"),
        ("mmr_esikKapasite", 1.0, "oran", "Kapasite politikası", "01.01.2026", "Aşım eşiği %100"),
        ("mmr_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("mmr_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("mmr_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("mmr_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("mmr_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("mmr_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("mmr_azamiSiparis", 100000, "adet", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("mmr_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("mmr_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("mmr_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("mmr_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("mmr_kararMetni", "PANO karar hücresi", "metin", "Karar motoru", "01.01.2026",
         "ÜRET/SATIN AL/AŞIM/VERİ YOK"),
        ("mmr_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("mmr_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("mmr_kuralMMR001", "Net ihtiyaç ≥ 0", "metin", "MMR-001", "01.01.2025",
         "Net ihtiyaç negatif olamaz"),
        ("mmr_kuralMMR002", "Kapasite yükü bayrağı", "metin", "MMR-002", "01.01.2025",
         "Yük > eşik → aşım"),
        ("mmr_kuralMMR003", "Brüt↔beklenen zincir", "metin", "MMR-003", "01.01.2026",
         "Zincir kırık görünür"),
        ("mmr_kuralMMR004", "Yetim mamul tespiti", "metin", "MMR-004", "01.01.2026",
         "Kart bağlanmalı"),
        ("mmr_kuralMMR005", "Karar yolu dört durum", "metin", "MMR-005", "01.01.2027",
         "ÜRET/SATIN AL/AŞIM/VERİ YOK"),
        ("mmr_hesapYillari", "2025-2027", "metin", "SPEC mevzuat", "01.01.2025",
         "Hesap yılları"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.15:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 33, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 34, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 35 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H34:L38", dh_bas)

    h(ws, 42, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 43, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 43, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    # HESAP satır eşlemesi: H01→C6 ... Hn→C(5+n)
    # H01 C6, H02 C7, H03 C8, H04 C9, H05 C10,
    # H06 C11, H07 C12, H08 C13, H09 C14, H10 C15,
    # H11 C16, H12 C17, H13 C18, H14 C19, H15 C20,
    # H16 C21, H17 C22, H18 C23, H19 C24, H20 C25,
    # H21 C26, H22 C27, H23 C28, H24 C29, H25 C30,
    # H26 C31, H27 C32, H28 C33, H29 C34, H30 C35,
    # H31 C36, H32 C37, H33 C38, H34 C39, H35 C40,
    # H36 C41, H37 C42, H38 C43, H39 C44, H40 C45,
    # H41 C46, H42 C47, H43 C48, H44 C49, H45 C50, H46 C51
    motor = [
        (44, "mmr_netIhtiyac", "=HESAP!C17"),
        (45, "mmr_onerilenSiparis", "=HESAP!C18"),
        (46, "mmr_kapasiteYuku", "=HESAP!C24"),
        (47, "mmr_kapasiteAsim", "=HESAP!C23"),
        (48, "mmr_satinAlAdet", "=HESAP!C25"),
        (49, "mmr_yetimKayit", "=HESAP!C19"),
        (50, "mmr_gecersizGecis", "=HESAP!C20"),
        (51, "mmr_zincirKirik", "=HESAP!C21"),
        (52, "mmr_ihtiyacSapma", "=HESAP!C28"),
        (53, "mmr_tahminAralik", "=HESAP!C39"),
        (54, "mmr_senaryoKarsilastirma", "=HESAP!C33"),
        (55, "mmr_tornadoZirve", "=HESAP!C34"),
        (56, "mmr_hhi", "=HESAP!C35"),
        (57, "mmr_kaliteSkor", "=HESAP!C36"),
        (58, "mmr_yuzdelikP90", "=HESAP!C40"),
        (59, "mmr_senaryoMotor", "=HESAP!C45"),
        (60, "mmr_kapaliDusum", "=HESAP!C16"),
        (61, "mmr_kuyrukOzeti", "=HESAP!C44"),
        (62, "mmr_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Net",HESAP!C17,"Yetim",HESAP!C19,'
         '"Zincir",HESAP!C21,"Aşım",HESAP!C23,"SatınAl",HESAP!C25)'),
        (63, "mmr_sonrakiKimlik",
         sonraki_kimlik_formul("MAM", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 65, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (66, "mmr_yorumNet",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Net ihtiyaç",mmr_netIhtiyac)'),
        (67, "mmr_yorumKapasite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ort kapasite yükü",TEXT(mmr_kapasiteYuku,"0.0%"))'),
        (68, "mmr_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"İhtiyaç sapması",mmr_ihtiyacSapma)'),
        (69, "mmr_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",mmr_tornadoZirve)'),
        (70, "mmr_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",mmr_senaryoKarsilastirma)'),
        (71, "mmr_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(mmr_hhi,"0.0%"))'),
        (72, "mmr_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",mmr_kaliteSkor)'),
        (73, "mmr_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",mmr_senaryoMotor)'),
        (74, "mmr_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",mmr_tahminAralik)'),
        (75, "mmr_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 brüt",mmr_yuzdelikP90)'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına mamul kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her BOM/iş emri satırını yazın; durum kolonunu listeden seçin.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "5. Kapalı durum (Tamam) açık kuyruklardan düşer.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma plan tarihine göredir.",
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
    for r in range(6, 31):
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
        "B4", CellIsRule(operator="equal", formula=['"ÜRET"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"SATIN AL"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"KAPASİTE AŞIMI"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 24):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["AKIS"]
    for col, val, fill, font in [
        ("S", '"YETİM"', kirmizi, kf),
        ("T", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("U", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("U", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("TEDARIK", sari), ("URETIM", kirmizi), ("TAMAM", yesil)]:
        ws.conditional_formatting.add(
            f"J{ILK}:J{SON}",
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
        "B7", CellIsRule(operator="equal", formula=['"ÜRET"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"SATIN AL"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"KAPASİTE AŞIMI"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(9, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KARTLAR"]
    ws.conditional_formatting.add(
        f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"D{ILK}:D{SON}",
        CellIsRule(operator="greaterThan", formula=["40"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "mmr_esikTedarikGun", "mmr_esikSatinAl",
        "mmr_esikKapasite",
        "mmr_senaryoTemkinli", "mmr_senaryoBaz", "mmr_senaryoIyimser",
        "mmr_tornadoOran", "mmr_yuzdelikOran", "mmr_olcekHedef", "mmr_azamiSiparis",
        "mmr_durumHaritasi", "mmr_kuyrukOzeti_metin", "mmr_kapaliDusum_metin",
        "mmr_kanitRaporu_metin", "mmr_kararMetni", "mmr_dosyaSurumu", "mmr_firmaUnvan",
        "mmr_kuralMMR001", "mmr_kuralMMR002", "mmr_kuralMMR003", "mmr_kuralMMR004",
        "mmr_kuralMMR005", "mmr_hesapYillari",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "mmr_netIhtiyac": 44, "mmr_onerilenSiparis": 45, "mmr_kapasiteYuku": 46,
        "mmr_kapasiteAsim": 47, "mmr_satinAlAdet": 48, "mmr_yetimKayit": 49,
        "mmr_gecersizGecis": 50, "mmr_zincirKirik": 51, "mmr_ihtiyacSapma": 52,
        "mmr_tahminAralik": 53, "mmr_senaryoKarsilastirma": 54, "mmr_tornadoZirve": 55,
        "mmr_hhi": 56, "mmr_kaliteSkor": 57, "mmr_yuzdelikP90": 58,
        "mmr_senaryoMotor": 59, "mmr_kapaliDusum": 60, "mmr_kuyrukOzeti": 61,
        "mmr_kanitRaporu": 62, "mmr_sonrakiKimlik": 63,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "mmr_yorumNet": 66, "mmr_yorumKapasite": 67, "mmr_yorumSapma": 68,
        "mmr_yorumTornado": 69, "mmr_yorumSenaryo": 70, "mmr_yorumHhi": 71,
        "mmr_yorumKalite": 72, "mmr_yorumSenMotor": 73, "mmr_yorumTahmin": 74,
        "mmr_yorumP90": 75,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
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
    dosya_ad = "MiniMrpBomMalzemeKapasite.xlsx"
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
