#!/usr/bin/env python3
"""Tahsilat Riski & Vade Komuta Paneli — A3 üretim betiği (manda v6)."""

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

URUN_AD = "Tahsilat Riski & Vade Komuta Paneli"
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
    {"durum_id": "ACIK", "durum_adi": "Açık", "sira": 1,
     "izinli_sonraki_durumlar": "HATIRLATMA", "kategori": "acik"},
    {"durum_id": "HATIRLATMA", "durum_adi": "Hatırlatma", "sira": 2,
     "izinli_sonraki_durumlar": "IHTAR|ACIK", "kategori": "acik"},
    {"durum_id": "IHTAR", "durum_adi": "İhtar", "sira": 3,
     "izinli_sonraki_durumlar": "TAHSIL|KAPALI", "kategori": "acik"},
    {"durum_id": "TAHSIL", "durum_adi": "Tahsil", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "KAPALI", "durum_adi": "Kapalı", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "ACIK|HATIRLATMA",
    "HATIRLATMA|IHTAR",
    "HATIRLATMA|ACIK",
    "IHTAR|TAHSIL",
    "IHTAR|KAPALI",
]

MUSTERILER = [
    ("MUS-00001", "Atlas Ticaret Ltd.", "Toptan", "İstanbul"),
    ("MUS-00002", "Demir Market A.Ş.", "Perakende", "Ankara"),
    ("MUS-00003", "Nova Gıda", "Gıda", "İzmir"),
    ("MUS-00004", "Proje Lojistik", "Lojistik", "Bursa"),
    ("MUS-00005", "Su Enerji Pro", "Enerji", "Antalya"),
    ("MUS-00006", "Cephe Yapı", "İnşaat", "Kocaeli"),
    ("MUS-00007", "Zemin Tarım", "Tarım", "Adana"),
    ("MUS-00008", "İskele Medikal", "Sağlık", "Gaziantep"),
]


def _yas_kova_trv(yas_hucre: str) -> str:
    """0-30 / 31-60 / 61-90 / 90+ — eşikler AYARLAR adlı aralıklarından."""
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=trv_kovaEsik1,"0-30",'
        f'IF({yas_hucre}<=trv_kovaEsik2,"31-60",'
        f'IF({yas_hucre}<=trv_kovaEsik3,"61-90","90+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Tahsilat kuyruğu ve kararı etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["ACIK", "HATIRLATMA", "IHTAR", "TAHSIL", "KAPALI",
                "HATIRLATMA", "IHTAR", "ACIK"]
    oncekiler = ["", "ACIK", "HATIRLATMA", "IHTAR", "IHTAR",
                 "ACIK", "HATIRLATMA", "ACIK"]
    for i in range(1, DEMO_AKIS + 1):
        kart = MUSTERILER[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "MUS-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: ACIK → IHTAR
            o, d = "ACIK", "IHTAR"
        if i == 20:  # geçersiz: IHTAR → HATIRLATMA
            o, d = "IHTAR", "HATIRLATMA"
        fatura = 80000 + i * 9500
        odeme = fatura if d == "TAHSIL" else 0
        if i in (12, 19):  # zincir kırık: ödeme ≠ fatura (kısmi/fazla)
            odeme = fatura - 12000 if d == "TAHSIL" else fatura + 45000
            if d not in ("TAHSIL", "KAPALI"):
                odeme = 0
                fatura_zincir = fatura + 45000  # fatura vs beklenen sapma için
            else:
                fatura_zincir = fatura
        else:
            fatura_zincir = fatura
        # Açık tutar: kapalıysa 0, değilse fatura
        tutar = 0 if d in ("TAHSIL", "KAPALI") else fatura
        vade = RAPOR_TARIHI - timedelta(days=(i * 4) % 110)
        odeme_tar = RAPOR_TARIHI - timedelta(days=2) if d == "TAHSIL" else None
        risk = min(95, 20 + (i * 3) % 70)
        satirlar.append({
            "id": f"FAT-{i:05d}",
            "kart": kart,
            "fatura_no": f"FR-2026-{i:04d}",
            "vade": vade,
            "odeme_tar": odeme_tar,
            "onceki": o,
            "durum": d,
            "fatura": fatura_zincir if i in (12, 19) and d not in ("TAHSIL", "KAPALI") else fatura,
            "odeme": odeme if d == "TAHSIL" else (odeme if i in (12, 19) and d == "TAHSIL" else 0),
            "tutar": tutar if i not in (12, 19) or d in ("TAHSIL", "KAPALI") else (
                fatura + 45000 if d not in ("TAHSIL", "KAPALI") else 0),
            "risk": risk,
            "sektor": MUSTERILER[(i - 1) % DEMO_KART][2],
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Takip notu {i}",
        })
    # Zincir için: fatura vs odeme tutarı — açık satırlarda odeme=0 olduğunda
    # zinciri fatura ile beklenen (tutar) üzerinden kuracağız.
    for s in satirlar:
        if s["id"] in ("FAT-00012", "FAT-00019") and s["durum"] not in ("TAHSIL", "KAPALI"):
            s["fatura"] = 80000 + int(s["id"][-5:]) * 9500
            s["odeme"] = s["fatura"] + 45000  # sapma görünür (ödeme alanı dolu ama durum açık)
            s["tutar"] = s["fatura"]
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Müşteri faturalarını vade yaşlandırması, tahsilat olasılığı ve durum makinesi "
      "ile yönetir; beklenen kaybı görünür kılar; ARA / İCRA HAZIRLIĞI / İZLE kararı üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Açık → Hatırlatma → İhtar → Tahsil / Kapalı",
        "Yaşlandırma kovaları (0-30 / 31-60 / 61-90 / 90+)",
        "Tahsilat olasılığı + beklenen kayıp motoru",
        "Geçersiz geçiş ve yetim müşteri uyarıları",
        "Dönemsel tahsilat kanıt raporu — yönetici/SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Tahsilat müdürü — günün kuyruğu ve icra hazırlığı",
        "CFO / mali işler — beklenen kayıp ve yaşlandırma",
        "SMMM — tahsilat kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a müşteri kartı girin. 2) AKIS'e fatura/vade satırlarını yazın. "
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
      'IF(OR(trv_yetimKayit>0,trv_gecersizGecis>0,trv_icraHazirlik>0),"İCRA HAZIRLIĞI",'
      'IF(trv_gecikenAdet>trv_esikGeciken,"ARA","İZLE")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: İZLE | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=trv_acikAdet", CATI),
        (7, "Hatırlatma Adet", '=COUNTIFS(tblAkis[Durum],"HATIRLATMA",tblAkis[Tutar],">0")', CATI),
        (8, "İhtar Adet", '=COUNTIFS(tblAkis[Durum],"IHTAR",tblAkis[Tutar],">0")', CATI),
        (9, "Geciken Adet", "=trv_gecikenAdet", CATI),
        (10, "Açık Tutar", "=trv_acikTutar", TL),
        (11, "Beklenen Kayıp", "=trv_beklenenKayip", TL),
        (12, "Yetim Kayıt", "=trv_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=trv_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=trv_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=trv_yasKova", None),
        (16, "Tahsilat Ort.", "=trv_tahsilatOrt", YÜZDE),
        (17, "Tahmin Aralık", "=trv_tahminAralik", TL),
        (18, "Senaryo Farkı", "=trv_senaryoKarsilastirma", TL),
        (19, "Tornado Zirve", "=trv_tornadoZirve", TL),
        (20, "HHI Yoğunlaşma", "=trv_hhi", YÜZDE),
        (21, "Kalite Skoru", "=trv_kaliteSkor", CATI),
        (22, "P90 Yüzdelik", "=trv_yuzdelikP90", TL),
        (23, "Kapalı Düşüm", "=trv_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=trv_yorumYas"),
        (8, "=trv_yorumTahsilatOrt"),
        (9, "=trv_yorumSapma"),
        (10, "=trv_yorumTornado"),
        (11, "=trv_yorumSenaryo"),
        (12, "=trv_yorumHhi"),
        (13, "=trv_yorumKalite"),
        (14, "=trv_yorumSenMotor"),
        (15, "=trv_yorumTahmin"),
        (16, "=trv_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Durum")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Açık", 6), ("Durum Hatırlatma", 8), ("Durum İhtar", 5),
        ("Durum Tahsil", 5), ("Durum Kapalı", 4),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Yaş Kova")
    h(ws, 26, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-30", 7), ("31-60", 8), ("61-90", 6), ("90+", 7),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Tahsilat")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 220000 + i * 18000, sayi=TL)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 3800000), ("Baz", 4500000), ("İyimser", 5200000),
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
    c3.title = "Haftalık Tahsilat"
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
    h(ws, 26, 17, "Kayıp")
    for i in range(6):
        h(ws, 27 + i, 16, i + 1, sayi=CATI)
        h(ws, 27 + i, 17, 120000 + i * 25000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = LineChart()
    c6.title = "Aylık Beklenen Kayıp"
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
    alt_bant(ws, 3, "Müşteri kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=trv_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: MUS-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Sektor", "Sehir", "LimitTl",
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
    for i, (kid, unvan, sektor, sehir) in enumerate(MUSTERILER):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="MUS-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Müşteri unvanı")
        _sari(ws, r, 3, sektor, baslik="Sektör", mesaj="Sektör")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, 1500000 + i * 120000, sayi=TL, baslik="Limit", mesaj="Risk limiti TL")
        _sari(ws, r, 6, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=TL if c == 5 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="MUS-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Sektör"), (4, "Şehir"), (6, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}",
              baslik="Limit", mesaj="Limit ≥ 0",
              hata_baslik="Limit", hata_mesaj="0 veya üzeri",
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
             "Fatura/vade satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "IslemId", "KaynakKartId", "FaturaNo", "VadeTarihi", "OdemeTarihi",
        "OncekiDurum", "Durum", "FaturaTutari", "OdemeTutari", "Tutar",
        "RiskSkoru", "Sektor", "YorumA", "YorumB",
        "GecikmeGun", "YasKova", "TahsilatOlasilik", "BeklenenKayip",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[VadeTarihi]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_trv("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[FaturaTutari]]",
        "tblAkis[[#This Row],[OdemeTutari]]",
        "zincirTolerans",
    ).lstrip("=")
    # Ödeme boş/0 ve açık durumdaysa zincir OK sayılmalı — ödeme 0 iken kırık olmasın
    zincir_guvenli = (
        f'IF(OR(tblAkis[[#This Row],[Durum]]="TAHSIL",'
        f'tblAkis[[#This Row],[OdemeTutari]]>0),{zincir_f},"OK")'
    )

    olasilik_f = (
        'IF(tblAkis[[#This Row],[YasKova]]="0-30",trv_olasilikK1,'
        'IF(tblAkis[[#This Row],[YasKova]]="31-60",trv_olasilikK2,'
        'IF(tblAkis[[#This Row],[YasKova]]="61-90",trv_olasilikK3,'
        'trv_olasilikK4)))'
    )

    form = {
        "GecikmeGun": (
            f'IF(tblAkis[[#This Row],[IslemId]]="","",'
            f'IF(OR(tblAkis[[#This Row],[Durum]]="TAHSIL",'
            f'tblAkis[[#This Row],[Durum]]="KAPALI"),0,{yas_f}))'
        ),
        "YasKova": f'IF(tblAkis[[#This Row],[IslemId]]="","",{kova_f})',
        "TahsilatOlasilik": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[Durum]]="TAHSIL",'
            f'tblAkis[[#This Row],[Durum]]="KAPALI"),"",{olasilik_f})'
        ),
        "BeklenenKayip": (
            'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            'tblAkis[[#This Row],[TahsilatOlasilik]]=""),"",'
            'tblAkis[[#This Row],[Tutar]]*(1-tblAkis[[#This Row],[TahsilatOlasilik]])'
            '*MAX(trv_riskTaban,tblAkis[[#This Row],[RiskSkoru]]/100))'
        ),
        "YetimBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[IslemId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[IslemId]]="","",{zincir_guvenli})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[IslemId]]="","",1)',
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="FAT-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["fatura_no"], baslik="Fatura No", mesaj="Fatura numarası")
        _sari(ws, r, 4, s["vade"], sayi=TARİH, baslik="Vade Tarihi", mesaj="Vade tarihi")
        _sari(ws, r, 5, s["odeme_tar"], sayi=TARİH, baslik="Ödeme Tarihi",
              mesaj="Ödeme tarihi (boş = ödenmedi)")
        _sari(ws, r, 6, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 7, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 8, s["fatura"], sayi=TL, baslik="Fatura Tutarı", mesaj="Fatura tutarı TL")
        _sari(ws, r, 9, s["odeme"], sayi=TL, baslik="Ödeme Tutarı", mesaj="Ödeme tutarı TL")
        _sari(ws, r, 10, s["tutar"], sayi=TL, baslik="Açık Tutar", mesaj="Açık bakiye TL")
        _sari(ws, r, 11, s["risk"], sayi=CATI, baslik="Risk Skoru", mesaj="0-100 risk skoru")
        _sari(ws, r, 12, s["sektor"], baslik="Sektör", mesaj="Sektör")
        _sari(ws, r, 13, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 14, s["yorum_b"], baslik="Yorum B", mesaj="Serbest takip notu")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 15):
            fmt = TARİH if c in (4, 5) else (
                TL if c in (8, 9, 10) else (CATI if c == 11 else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:V{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="FAT-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "textLength", "0", f"C{ILK}:C{SON}",
              baslik="Fatura No", mesaj="Fatura numarası",
              hata_baslik="Fatura", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Vade", mesaj="Vade tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "date", "1", f"E{ILK}:E{SON}",
              baslik="Ödeme", mesaj="Ödeme tarihi (opsiyonel)",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"G{ILK}:G{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(8, "Fatura"), (9, "Ödeme"), (10, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "whole", "0", f"K{ILK}:K{SON}",
              baslik="Risk", mesaj="0-100",
              hata_baslik="Risk", hata_mesaj="0-100 arası",
              isaret="between", f2="100")
    dogrulama(ws, "list", "ListeYorumAB", f"M{ILK}:M{SON}",
              baslik="Yorum A", mesaj="A veya B",
              hata_baslik="Liste", hata_mesaj="A veya B")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 23)})
    ws.column_dimensions["N"].width = 16
    ws.column_dimensions["T"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # Explicit Durum/Tutar (büyük harf) — şablon yardımcıları küçük harf kullanır
    h(ws, 6, 1, "Kuyruk Açık")
    h(ws, 6, 2,
      '=COUNTIFS(tblAkis[Durum],"ACIK",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 6, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ACIK")', sayi=TL)

    h(ws, 7, 1, "Kuyruk Hatırlatma")
    h(ws, 7, 2,
      '=COUNTIFS(tblAkis[Durum],"HATIRLATMA",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 7, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"HATIRLATMA")', sayi=TL)

    h(ws, 8, 1, "Kuyruk İhtar")
    h(ws, 8, 2,
      '=COUNTIFS(tblAkis[Durum],"IHTAR",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 8, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"IHTAR")', sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[GecikmeGun],">"&trv_esikGecikmeGun,tblAkis[Durum],"<>TAHSIL",'
      'tblAkis[Durum],"<>KAPALI",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&trv_esikGecikmeGun,'
      'tblAkis[Durum],"<>TAHSIL",tblAkis[Durum],"<>KAPALI")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Tahsil+Kapalı) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Tahsil")
    h(ws, 12, 2, '=COUNTIF(tblAkis[Durum],"TAHSIL")', sayi=CATI)
    h(ws, 12, 3, '=SUMIF(tblAkis[Durum],"TAHSIL",tblAkis[OdemeTutari])', sayi=TL)
    h(ws, 13, 1, "Kapalı")
    h(ws, 13, 2, '=COUNTIF(tblAkis[Durum],"KAPALI")', sayi=CATI)
    h(ws, 13, 3, '=SUMIF(tblAkis[Durum],"KAPALI",tblAkis[Tutar])', sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=trv_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-30", "31-60", "61-90", "90+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[Tutar])', sayi=TL)

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
        ("H03", "Toplam açık tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "Fatura toplam", "=SUM(tblAkis[FaturaTutari])"),
        ("H05", "Ödeme toplam", "=SUM(tblAkis[OdemeTutari])"),
        ("H06", "Açık adet", '=COUNTIF(tblAkis[Durum],"ACIK")'),
        ("H07", "Hatırlatma adet", '=COUNTIF(tblAkis[Durum],"HATIRLATMA")'),
        ("H08", "İhtar adet", '=COUNTIF(tblAkis[Durum],"IHTAR")'),
        ("H09", "Tahsil adet", '=COUNTIF(tblAkis[Durum],"TAHSIL")'),
        ("H10", "Kapalı adet", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        ("H11", "Açık kuyruk adet", "=C11+C12+C13"),
        ("H12", "Kapalı düşüm adet", "=C14+C15"),
        ("H13", "Açık tutar kuyruk",
         '=SUMIF(tblAkis[Durum],"ACIK",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"HATIRLATMA",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"IHTAR",tblAkis[Tutar])'),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Yaş 0-30", '=COUNTIF(tblAkis[YasKova],"0-30")'),
        ("H19", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H20", "Yaş 61-90", '=COUNTIF(tblAkis[YasKova],"61-90")'),
        ("H21", "Yaş 90+", '=COUNTIF(tblAkis[YasKova],"90+")'),
        ("H22", "Ort gecikme gün",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[GecikmeGun]),0)"),
        ("H23", "Geciken adet",
         '=COUNTIFS(tblAkis[GecikmeGun],">"&trv_esikGecikmeGun,tblAkis[Durum],"<>TAHSIL",'
         'tblAkis[Durum],"<>KAPALI",tblAkis[Tutar],">0")'),
        ("H24", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[GecikmeGun],">"&trv_esikGecikmeGun,'
         'tblAkis[Durum],"<>TAHSIL",tblAkis[Durum],"<>KAPALI")'),
        ("H25", "Beklenen kayıp",
         "=IFERROR(SUMIF(tblAkis[KayitDolu],1,tblAkis[BeklenenKayip]),0)"),
        ("H26", "İcra hazırlık adet",
         '=COUNTIFS(tblAkis[Durum],"IHTAR",tblAkis[Tutar],">0")'
         '+COUNTIFS(tblAkis[YasKova],"90+",tblAkis[Durum],"<>TAHSIL",'
         'tblAkis[Durum],"<>KAPALI")'),
        ("H27", "Tahsilat ort olasılık",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[TahsilatOlasilik]),0)"),
        ("H28", "Tahsilat sapma", "=C9-C10"),
        ("H29", "Temkinli senaryo", "=C18*trv_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C18*trv_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C18*trv_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×tutar)", "=C18*trv_tornadoOran"),
        ("H34", "HHI proxy",
         "=IFERROR((MAX(C11,C12,C13)/MAX(C16,1))^2,0)"),
        ("H35", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,100-C19*10-C20*10-C21*10-C28*2))"),
        ("H36", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],trv_yuzdelikOran),0)"),
        ("H37", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[FaturaTutari],trv_yuzdelikOran),0)"),
        ("H40", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H41", "Limit toplam", "=SUM(tblKartlar[LimitTl])"),
        ("H42", "Limit kullanım oranı", "=IFERROR(C18/MAX(C46,1),0)"),
        ("H43", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",C11,"Hatırlatma",C12,"İhtar",C13,"Geciken",C28)'),
        ("H44", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C23,C24,C25,C26)'),
        ("H45", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C34,"Baz",C35,"İyimser",C36)'),
        ("H46", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "kayıp", "senaryo", "tornado", "tahmin", "limit", "fark",
            "toplam", "ödeme", "fatura", "yüzdelik", "p90")) and "oran" not in acik.lower() and "adet" not in acik.lower() and "metin" not in acik.lower() and "özet" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "olasılık" in acik.lower() or "hhi" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "gün" in acik.lower():
            fmt = GUN
        if "kalite" in acik.lower():
            fmt = CATI
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 55, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 56, 1, "Temkinli")
    h(ws, 56, 2, "=trv_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 56, 3, "=C34", sayi=TL)
    h(ws, 57, 1, "Baz")
    h(ws, 57, 2, "=trv_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 57, 3, "=C35", sayi=TL)
    h(ws, 58, 1, "İyimser")
    h(ws, 58, 2, "=trv_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 58, 3, "=C36", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Tahsilat Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=trv_acikAdet", CATI),
        (10, "Açık Tutar", "=trv_acikTutar", TL),
        (11, "Beklenen Kayıp", "=trv_beklenenKayip", TL),
        (12, "Yetim", "=trv_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=trv_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=trv_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=trv_tahminAralik", TL),
        (16, "Kanıt Özeti", "=trv_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Tahsilat olasılıkları yaş kovasına göre AYARLAR'dan gelir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 23, 1, "2. İhtar ve 90+ kovadaki müşteriler için icra hazırlığını başlatın.")
    h(ws, 24, 1, "3. Zincir kırık satırlarda fatura/ödeme farkını açıklayın.")
    h(ws, 26, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H26", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Açık", '=COUNTIF(tblAkis[Durum],"ACIK")'),
        (8, "Hatırlatma", '=COUNTIF(tblAkis[Durum],"HATIRLATMA")'),
        (9, "İhtar", '=COUNTIF(tblAkis[Durum],"IHTAR")'),
        (10, "Tahsil", '=COUNTIF(tblAkis[Durum],"TAHSIL")'),
        (11, "Kapalı", '=COUNTIF(tblAkis[Durum],"KAPALI")'),
        (12, "Açık toplam", "=B7+B8+B9"),
        (13, "Yetim", "=trv_yetimKayit"),
        (14, "Geçersiz", "=trv_gecersizGecis"),
        (15, "Zincir kırık", "=trv_zincirKirik"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=trv_kaliteSkor"),
        (18, "Açık tutar", "=trv_acikTutar"),
        (19, "Beklenen kayıp", "=trv_beklenenKayip"),
        (20, "Kapalı düşüm", "=trv_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r >= 18 else (CATI if r < 16 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} müşteri kartı ve {DEMO_AKIS} fatura satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Açık/Hatırlatma/İhtar/Tahsil/Kapalı — kapalılar kuyruktan düşer")
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

    h(ws, 3, 9, "Sektör", kalin=True)
    h(ws, 6, 9, "Sektor", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Toptan", "Perakende", "Gıda", "Lojistik", "Enerji", "İnşaat"], 7):
        _sari(ws, i, 9, b, baslik="Sektör", mesaj="Sektör listesi")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 10000, "TL", "İş kuralı — fatura/ödeme sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("trv_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("trv_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("trv_kovaEsik1", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("trv_kovaEsik2", 60, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("trv_kovaEsik3", 90, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("trv_olasilikK1", 0.95, "oran", "Tahsilat modeli", "01.01.2026", "0-30 gün olasılık"),
        ("trv_olasilikK2", 0.75, "oran", "Tahsilat modeli", "01.01.2026", "31-60 gün olasılık"),
        ("trv_olasilikK3", 0.50, "oran", "Tahsilat modeli", "01.01.2026", "61-90 gün olasılık"),
        ("trv_olasilikK4", 0.25, "oran", "Tahsilat modeli", "01.01.2026", "90+ gün olasılık"),
        ("trv_riskTaban", 0.5, "oran", "Risk modeli", "01.01.2026", "Risk çarpan tabanı"),
        ("trv_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("trv_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("trv_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("trv_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("trv_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("trv_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("trv_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("trv_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("trv_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("trv_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("trv_kanitRaporu_metin", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("trv_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("trv_firmaUnvan", "Örnek Ticaret A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
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
    tablo_ekle(ws, "tblDurumHaritasi", "H34:L39", dh_bas)

    h(ws, 42, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 43, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 43, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (44, "trv_acikAdet", "=HESAP!C16"),
        (45, "trv_acikTutar", "=HESAP!C18"),
        (46, "trv_gecikenAdet", "=HESAP!C28"),
        (47, "trv_beklenenKayip", "=HESAP!C30"),
        (48, "trv_yetimKayit", "=HESAP!C19"),
        (49, "trv_gecersizGecis", "=HESAP!C20"),
        (50, "trv_zincirKirik", "=HESAP!C21"),
        (51, "trv_icraHazirlik", "=HESAP!C31"),
        (52, "trv_yasKova", "=HESAP!C49"),
        (53, "trv_tahsilatOrt", "=HESAP!C32"),
        (54, "trv_tahsilatSapma", "=HESAP!C33"),
        (55, "trv_tahminAralik", "=HESAP!C43"),
        (56, "trv_senaryoKarsilastirma", "=HESAP!C37"),
        (57, "trv_tornadoZirve", "=HESAP!C38"),
        (58, "trv_hhi", "=HESAP!C39"),
        (59, "trv_kaliteSkor", "=HESAP!C40"),
        (60, "trv_yuzdelikP90", "=HESAP!C44"),
        (61, "trv_senaryoMotor", "=HESAP!C50"),
        (62, "trv_kapaliDusum", "=HESAP!C17"),
        (63, "trv_kuyrukOzeti", "=HESAP!C48"),
        (64, "trv_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C16,"Yetim",HESAP!C19,'
         '"Zincir",HESAP!C21,"Geciken",HESAP!C28,"Kayıp",HESAP!C30)'),
        (65, "trv_sonrakiKimlik",
         sonraki_kimlik_formul("MUS", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 67, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (68, "trv_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",trv_yasKova)'),
        (69, "trv_yorumTahsilatOrt",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ort tahsilat olasılığı",TEXT(trv_tahsilatOrt,"0.0%"))'),
        (70, "trv_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Fatura-ödeme sapması",TEXT(trv_tahsilatSapma,"0.00"),"TL")'),
        (71, "trv_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(trv_tornadoZirve,"0.00"),"TL")'),
        (72, "trv_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(trv_senaryoKarsilastirma,"0.00"),"TL")'),
        (73, "trv_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(trv_hhi,"0.0%"))'),
        (74, "trv_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",trv_kaliteSkor)'),
        (75, "trv_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",trv_senaryoMotor)'),
        (76, "trv_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(trv_tahminAralik,"0.00"),"TL")'),
        (77, "trv_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 fatura",TEXT(trv_yuzdelikP90,"0.00"),"TL")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına müşteri kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her fatura/vade satırını yazın; durum kolonunu listeden seçin.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "5. Kapalı durumlar (Tahsil/Kapalı) açık kuyruklardan düşer.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma vade tarihine göredir.",
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
    for r in range(7, 12):
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
        "B4", CellIsRule(operator="equal", formula=['"İZLE"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"ARA"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"İCRA HAZIRLIĞI"'], fill=kirmizi, font=kf))
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
    for durum, fill in [("HATIRLATMA", sari), ("IHTAR", kirmizi), ("TAHSIL", yesil)]:
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
        "B7", CellIsRule(operator="equal", formula=['"İZLE"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"ARA"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"İCRA HAZIRLIĞI"'], fill=kirmizi, font=kf))
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
        CellIsRule(operator="greaterThan", formula=["3000000"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "trv_esikGecikmeGun", "trv_esikGeciken",
        "trv_kovaEsik1", "trv_kovaEsik2", "trv_kovaEsik3",
        "trv_olasilikK1", "trv_olasilikK2", "trv_olasilikK3", "trv_olasilikK4",
        "trv_riskTaban",
        "trv_senaryoTemkinli", "trv_senaryoBaz", "trv_senaryoIyimser",
        "trv_tornadoOran", "trv_yuzdelikOran", "trv_olcekHedef", "trv_azamiTutar",
        "trv_durumHaritasi", "trv_kuyrukOzeti_metin", "trv_kapaliDusum_metin",
        "trv_kanitRaporu_metin", "trv_dosyaSurumu", "trv_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "trv_acikAdet": 44, "trv_acikTutar": 45, "trv_gecikenAdet": 46,
        "trv_beklenenKayip": 47, "trv_yetimKayit": 48, "trv_gecersizGecis": 49,
        "trv_zincirKirik": 50, "trv_icraHazirlik": 51, "trv_yasKova": 52,
        "trv_tahsilatOrt": 53, "trv_tahsilatSapma": 54, "trv_tahminAralik": 55,
        "trv_senaryoKarsilastirma": 56, "trv_tornadoZirve": 57, "trv_hhi": 58,
        "trv_kaliteSkor": 59, "trv_yuzdelikP90": 60, "trv_senaryoMotor": 61,
        "trv_kapaliDusum": 62, "trv_kuyrukOzeti": 63, "trv_kanitRaporu": 64,
        "trv_sonrakiKimlik": 65,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "trv_yorumYas": 68, "trv_yorumTahsilatOrt": 69, "trv_yorumSapma": 70,
        "trv_yorumTornado": 71, "trv_yorumSenaryo": 72, "trv_yorumHhi": 73,
        "trv_yorumKalite": 74, "trv_yorumSenMotor": 75, "trv_yorumTahmin": 76,
        "trv_yorumP90": 77,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
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
    dosya_ad = "TahsilatRiskiVadeKomutaPaneli.xlsx"
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
