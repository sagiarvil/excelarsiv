#!/usr/bin/env python3
"""Makine Bakım & Kalibrasyon + Duruş Maliyeti — A3+A4 üretim betiği (manda v6)."""

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
from ortak.kanit_uretici import (  # noqa: E402
    imza_alani,
    rapor_baslik_satirlari,
    sha_placeholder_notu,
)

URUN_AD = "Makine Bakım & Kalibrasyon + Duruş Maliyeti"
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
     "izinli_sonraki_durumlar": "YAPILDI|GECIKME", "kategori": "acik"},
    {"durum_id": "YAPILDI", "durum_adi": "Yapıldı", "sira": 2,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "GECIKME", "durum_adi": "Gecikme", "sira": 3,
     "izinli_sonraki_durumlar": "YAPILDI|KRITIK", "kategori": "acik"},
    {"durum_id": "KRITIK", "durum_adi": "Kritik", "sira": 4,
     "izinli_sonraki_durumlar": "YAPILDI", "kategori": "acik"},
])

IZINLI_GECIS = [
    "PLAN|YAPILDI",
    "PLAN|GECIKME",
    "GECIKME|YAPILDI",
    "GECIKME|KRITIK",
    "KRITIK|YAPILDI",
]

MAKINELER = [
    ("MAK-00001", "CNC Torna T-100", "Hat-A", "CNC"),
    ("MAK-00002", "Pres Hidrolik P-50", "Hat-A", "Pres"),
    ("MAK-00003", "Kaynak Robot KR-20", "Hat-B", "Robot"),
    ("MAK-00004", "Freze Merkezi FM-30", "Hat-B", "CNC"),
    ("MAK-00005", "Konveyör K-08", "Hat-C", "Konveyor"),
    ("MAK-00006", "Kompresör KMP-12", "Hat-C", "Kompresor"),
    ("MAK-00007", "Ölçüm Tezgâhı OT-05", "Kalite", "Olcu"),
    ("MAK-00008", "Fırın Isıl F-40", "Hat-A", "Firin"),
]

BAKIM_TURLERI = ["ONLEYICI", "DUZELTICI", "KALIBRASYON", "ACIL"]


def _yas_kova_mbk(yas_hucre: str) -> str:
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=mbk_kovaEsik1,"0-7",'
        f'IF({yas_hucre}<=mbk_kovaEsik2,"8-30",'
        f'IF({yas_hucre}<=mbk_kovaEsik3,"31-60","60+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Bakım planı ve duruş maliyeti kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    satirlar = []
    durumlar = ["PLAN", "GECIKME", "YAPILDI", "KRITIK",
                "GECIKME", "PLAN", "YAPILDI", "PLAN"]
    oncekiler = ["", "PLAN", "GECIKME", "GECIKME",
                 "PLAN", "PLAN", "KRITIK", "PLAN"]
    for i in range(1, DEMO_AKIS + 1):
        kart = MAKINELER[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "MAK-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "PLAN", "KRITIK"
        if i == 20:
            o, d = "YAPILDI", "PLAN"
        plan_durus = 2.0 + (i % 4) * 0.5
        fiili = plan_durus
        if i in (12, 19):
            fiili = plan_durus + 1.5
        if i == 15:
            fiili = 6.0
            plan_durus = 2.0
        plan_tar = RAPOR_TARIHI - timedelta(days=(i * 2) % 45)
        gercek = plan_tar + timedelta(days=(0 if d == "YAPILDI" else (i % 5)))
        kal_kalan = 30 - (i * 3) % 50
        satirlar.append({
            "id": f"BKM-{i:05d}",
            "kart": kart,
            "plan_tar": plan_tar,
            "gercek": gercek if d == "YAPILDI" else None,
            "plan_durus": plan_durus,
            "fiili": fiili,
            "kal_kalan": kal_kalan,
            "tur": BAKIM_TURLERI[i % len(BAKIM_TURLERI)],
            "onceki": o,
            "durum": d,
            "imza": f"İMZA-{i:03d}" if d == "YAPILDI" else "",
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Bakım notu {i}",
            "saatlik": 850 + (i % 5) * 50,
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Makine bakım planı, kalibrasyon vadesi ve duruş saatini maliyetle bağlar; "
      "durum makinesi ve imza kanıt zinciri üretir; PLANLI / GECİKMİŞ / KRİTİK / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Plan → Yapıldı / Gecikme → Kritik",
        "Günlük/haftalık duruş + saatlik duruş maliyeti hesabı",
        "Plan tarih ↔ gerçekleşen ↔ duruş saati zincir bayrağı",
        "Geçersiz geçiş ve yetim makine uyarıları",
        "KANIT_RAPORU — denetçi/SMMM imza alanlı dönem kanıtı",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Bakım / kalibrasyon — plan kapanışı ve vade takibi",
        "Üretim müdürü — PLANLI / GECİKMİŞ / KRİTİK kararı",
        "Denetçi / SMMM — KANIT_RAPORU dosyalama",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) MAKINELER'e kart girin. 2) AKIS'e bakım/kalibrasyon satırlarını yazın. "
      "3) KUYRUKLAR ve PANO'dan kararı görün. 4) KANIT_RAPORU'ndan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def makineler(ws):
    sayfa_hazirla(ws, "MAKINELER", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Makine kartları — her varlığın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=mbk_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: MAK-00009 | Risk: Elle çakışma")

    sutunlar = ["MakineId", "Ad", "Hat", "Tip", "SaatlikDurusMaliyeti",
                "KalibrasyonPeriyotGun", "Aktif", "ToplamDurus", "OrtDurus"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamDurus": (
            'IF(tblMakineler[[#This Row],[MakineId]]="","",'
            'SUMIF(tblAkis[KaynakMakineId],tblMakineler[[#This Row],[MakineId]],'
            'tblAkis[FiiliDurusSaat]))'
        ),
        "OrtDurus": (
            'IF(tblMakineler[[#This Row],[MakineId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakMakineId],'
            'tblMakineler[[#This Row],[MakineId]],tblAkis[FiiliDurusSaat]),0))'
        ),
    }
    for i, (kid, ad, hat, tip) in enumerate(MAKINELER):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Makine Kimliği", mesaj="MAK-##### formatında")
        _sari(ws, r, 2, ad, baslik="Makine Adı", mesaj="Makine adı")
        _sari(ws, r, 3, hat, baslik="Hat", mesaj="Üretim hattı")
        _sari(ws, r, 4, tip, baslik="Tip", mesaj="Makine tipi")
        _sari(ws, r, 5, 850 + i * 50, sayi=TL, baslik="Saatlik Duruş",
              mesaj="Saatlik duruş maliyeti TL")
        _sari(ws, r, 6, 90 + i * 15, sayi=CATI, baslik="Kalibrasyon Periyot",
              mesaj="Kalibrasyon periyodu (gün)")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=TL if c == 5 else (CATI if c == 6 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblMakineler", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Makine Kimliği", mesaj="MAK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Ad"), (3, "Hat"), (4, "Tip")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}",
              baslik="Saatlik Duruş", mesaj="TL ≥ 0",
              hata_baslik="Maliyet", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Kalibrasyon Periyot", mesaj="Gün ≥ 0",
              hata_baslik="Periyot", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 24


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=26)
    alt_bant(ws, 3,
             "Bakım satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "BakimId", "KaynakMakineId", "PlanTarih", "GerceklesenTarih",
        "PlanDurusSaat", "FiiliDurusSaat", "KalibrasyonGunKalan", "BakimTuru",
        "OncekiDurum", "Durum", "ImzaKanit", "YorumA", "YorumB", "SaatlikMaliyet",
        "HesaplananDurus", "GecikmeGun", "YasKova",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu", "DurusMaliyet",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[PlanTarih]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_mbk("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakMakineId]]",
        "tblMakineler[MakineId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[PlanDurusSaat]]",
        "tblAkis[[#This Row],[FiiliDurusSaat]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "HesaplananDurus": (
            'IF(tblAkis[[#This Row],[BakimId]]="","",'
            'tblAkis[[#This Row],[FiiliDurusSaat]])'
        ),
        "GecikmeGun": (
            f'IF(tblAkis[[#This Row],[BakimId]]="","",'
            f'IF(tblAkis[[#This Row],[Durum]]="YAPILDI",0,{yas_f}))'
        ),
        "YasKova": f'IF(tblAkis[[#This Row],[BakimId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[BakimId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[BakimId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[BakimId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[BakimId]]="","",1)',
        "DurusMaliyet": (
            'IF(tblAkis[[#This Row],[BakimId]]="","",'
            'tblAkis[[#This Row],[FiiliDurusSaat]]*'
            'tblAkis[[#This Row],[SaatlikMaliyet]]*mbk_durusKatsayi)'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Bakım Kimliği", mesaj="BKM-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Makine",
              mesaj="MAKINELER'deki MakineId")
        _sari(ws, r, 3, s["plan_tar"], sayi=TARİH, baslik="Plan Tarih",
              mesaj="Planlanan bakım tarihi")
        _sari(ws, r, 4, s["gercek"], sayi=TARİH, baslik="Gerçekleşen",
              mesaj="Gerçekleşen bakım tarihi")
        _sari(ws, r, 5, s["plan_durus"], sayi=CATI, baslik="Plan Duruş",
              mesaj="Planlanan duruş saati")
        _sari(ws, r, 6, s["fiili"], sayi=CATI, baslik="Fiili Duruş",
              mesaj="Gerçekleşen duruş saati")
        _sari(ws, r, 7, s["kal_kalan"], sayi=CATI, baslik="Kalibrasyon Kalan",
              mesaj="Kalibrasyona kalan gün (negatif = aşım)")
        _sari(ws, r, 8, s["tur"], baslik="Bakım Türü", mesaj="Bakım türü listesinden")
        _sari(ws, r, 9, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 10, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 11, s["imza"] or None, baslik="İmza/Kanıt", mesaj="İmza veya kanıt notu")
        _sari(ws, r, 12, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 13, s["yorum_b"], baslik="Yorum B", mesaj="Serbest not")
        _sari(ws, r, 14, s["saatlik"], sayi=TL, baslik="Saatlik Maliyet",
              mesaj="Saatlik duruş maliyeti TL")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 15):
            fmt = TARİH if c in (3, 4) else (
                TL if c == 14 else (CATI if c in (5, 6, 7) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:V{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Bakım Kimliği", mesaj="BKM-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Makine", mesaj="Makine kimliği (boş = yetim riski)",
              hata_baslik="Makine", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Plan Tarih", mesaj="Planlanan bakım tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Gerçekleşen", mesaj="Gerçekleşen tarih (boş olabilir)",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(5, "Plan Duruş"), (6, "Fiili Duruş")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "whole", "-365", f"G{ILK}:G{SON}",
              baslik="Kalibrasyon Kalan", mesaj="Gün (negatif aşım)",
              hata_baslik="Gün", hata_mesaj="-365 ile 3650 arası",
              isaret="between", f2="3650")
    dogrulama(ws, "list", "ListeBakimTuru", f"H{ILK}:H{SON}",
              baslik="Bakım Türü", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"I{ILK}:I{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"J{ILK}:J{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeYorumAB", f"L{ILK}:L{SON}",
              baslik="Yorum A", mesaj="A veya B",
              hata_baslik="Liste", hata_mesaj="A veya B")
    dogrulama(ws, "decimal", "0", f"N{ILK}:N{SON}",
              baslik="Saatlik Maliyet", mesaj="TL ≥ 0",
              hata_baslik="Maliyet", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 23)})
    ws.column_dimensions["M"].width = 16
    ws.column_dimensions["S"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Duruş Saat", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Kuyruk Plan")
    h(ws, 6, 2, '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 6, 3, '=SUMIF(tblAkis[Durum],"PLAN",tblAkis[FiiliDurusSaat])', sayi=CATI)

    h(ws, 7, 1, "Kuyruk Gecikme")
    h(ws, 7, 2, '=COUNTIFS(tblAkis[Durum],"GECIKME",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 7, 3, '=SUMIF(tblAkis[Durum],"GECIKME",tblAkis[FiiliDurusSaat])', sayi=CATI)

    h(ws, 8, 1, "Kuyruk Kritik")
    h(ws, 8, 2, '=COUNTIFS(tblAkis[Durum],"KRITIK",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 8, 3, '=SUMIF(tblAkis[Durum],"KRITIK",tblAkis[FiiliDurusSaat])', sayi=CATI)

    h(ws, 9, 1, "Duruş eşik aşımı (fiili>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[FiiliDurusSaat],">"&mbk_esikDurusSaat,tblAkis[Durum],"<>YAPILDI",'
      'tblAkis[KayitDolu],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[FiiliDurusSaat],tblAkis[FiiliDurusSaat],">"&mbk_esikDurusSaat,'
      'tblAkis[Durum],"<>YAPILDI")',
      sayi=CATI)

    h(ws, 11, 1, "Kapalı (Yapıldı) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Yapıldı")
    h(ws, 12, 2, '=COUNTIF(tblAkis[Durum],"YAPILDI")', sayi=CATI)
    h(ws, 12, 3, '=SUMIF(tblAkis[Durum],"YAPILDI",tblAkis[FiiliDurusSaat])', sayi=CATI)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=mbk_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[FiiliDurusSaat])', sayi=CATI)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam bakım", '=COUNTA(tblAkis[BakimId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Toplam duruş saat", "=SUM(tblAkis[FiiliDurusSaat])"),
        ("H04", "Hesaplanan duruş toplam", "=SUM(tblAkis[HesaplananDurus])"),
        ("H05", "Plan duruş toplam", "=SUM(tblAkis[PlanDurusSaat])"),
        ("H06", "Plan adet", '=COUNTIF(tblAkis[Durum],"PLAN")'),
        ("H07", "Gecikme adet", '=COUNTIF(tblAkis[Durum],"GECIKME")'),
        ("H08", "Kritik adet", '=COUNTIF(tblAkis[Durum],"KRITIK")'),
        ("H09", "Yapıldı adet", '=COUNTIF(tblAkis[Durum],"YAPILDI")'),
        ("H10", "Kalibrasyon aşım adet",
         '=COUNTIFS(tblAkis[KalibrasyonGunKalan],"<0",tblAkis[KayitDolu],">0")'),
        ("H11", "Açık kuyruk adet", "=C11+C12+C13"),
        ("H12", "Kapalı düşüm adet", "=C14"),
        ("H13", "Açık duruş kuyruk",
         '=SUMIF(tblAkis[Durum],"PLAN",tblAkis[FiiliDurusSaat])'
         '+SUMIF(tblAkis[Durum],"GECIKME",tblAkis[FiiliDurusSaat])'
         '+SUMIF(tblAkis[Durum],"KRITIK",tblAkis[FiiliDurusSaat])'),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H19", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H20", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H21", "Yaş 60+", '=COUNTIF(tblAkis[YasKova],"60+")'),
        ("H22", "Ort duruş saat",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[FiiliDurusSaat]),0)"),
        ("H23", "Duruş eşik aşım adet",
         '=COUNTIFS(tblAkis[FiiliDurusSaat],">"&mbk_esikDurusSaat,tblAkis[Durum],"<>YAPILDI",'
         'tblAkis[KayitDolu],">0")'),
        ("H24", "Duruş eşik aşım saat",
         '=SUMIFS(tblAkis[FiiliDurusSaat],tblAkis[FiiliDurusSaat],">"&mbk_esikDurusSaat,'
         'tblAkis[Durum],"<>YAPILDI")'),
        ("H25", "Duruş maliyet toplam",
         "=IFERROR(SUMIF(tblAkis[KayitDolu],1,tblAkis[DurusMaliyet]),0)"),
        ("H26", "Kalibrasyon aşım proxy", "=C15"),
        ("H27", "Haftalık duruş proxy", "=C8*mbk_haftalikCarpan"),
        ("H28", "Duruş sapma (fiili-plan)", "=C8-C10"),
        ("H29", "Temkinli senaryo", "=C18*mbk_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C18*mbk_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C18*mbk_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×duruş)", "=C18*mbk_tornadoOran"),
        ("H34", "HHI proxy",
         "=IFERROR((MAX(C11,C12,C13)/MAX(C16,1))^2,0)"),
        ("H35", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,mbk_kaliteTaban-C19*mbk_kaliteCeza-C20*mbk_kaliteCeza"
         "-C21*mbk_kaliteCeza-C28*mbk_kaliteSapma))"),
        ("H36", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[FiiliDurusSaat],mbk_yuzdelikOran),0)"),
        ("H37", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[BakimId])+1,tblAkis[FiiliDurusSaat],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[FiiliDurusSaat],mbk_yuzdelikOran),0)"),
        ("H40", "Makine sayısı", '=COUNTA(tblMakineler[MakineId])'),
        ("H41", "Periyot toplam", "=SUM(tblMakineler[KalibrasyonPeriyotGun])"),
        ("H42", "Duruş / makine oranı", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H43", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Plan",C11,"Gecikme",C12,"Kritik",C13,"Eşik",C28)'),
        ("H44", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C23,C24,C25,C26)'),
        ("H45", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C34,"Baz",C35,"İyimser",C36)'),
        ("H46", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H47", "Gecikmiş karar adet",
         '=IF(OR(C12>0,C28>0),1,0)'),
        ("H48", "Kritik karar adet",
         '=IF(OR(C13>0,C15>0),1,0)'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "maliyet", "senaryo", "tornado", "tahmin", "yüzdelik", "p90", "fark",
        )) and "oran" not in acik.lower() and "adet" not in acik.lower() and "metin" not in acik.lower() and "özet" not in acik.lower() and "saat" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "hhi" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "gün" in acik.lower() and "adet" not in acik.lower():
            fmt = GUN
        if "kalite" in acik.lower():
            fmt = CATI
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 57, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 57, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 57, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 58, 1, "Temkinli")
    h(ws, 58, 2, "=IFERROR(C34/MAX(C18,1),0)", sayi=YÜZDE)
    h(ws, 58, 3, "=C34", sayi=CATI)
    h(ws, 59, 1, "Baz")
    h(ws, 59, 2, "=IFERROR(C35/MAX(C18,1),0)", sayi=YÜZDE)
    h(ws, 59, 3, "=C35", sayi=CATI)
    h(ws, 60, 1, "İyimser")
    h(ws, 60, 2, "=IFERROR(C36/MAX(C18,1),0)", sayi=YÜZDE)
    h(ws, 60, 3, "=C36", sayi=CATI)

    genislik(ws, {"A": 10, "B": 32, "C": 50})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblAkis[BakimId])=0,"VERİ YOK",'
      'IF(mbk_kritikAdet>0,"KRİTİK",'
      'IF(OR(mbk_gecikmeAdet>0,mbk_durusEsikAsim>0),"GECİKMİŞ","PLANLI")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: PLANLI | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=mbk_acikAdet", CATI),
        (7, "Plan Adet", '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[KayitDolu],">0")', CATI),
        (8, "Gecikme Adet", '=COUNTIFS(tblAkis[Durum],"GECIKME",tblAkis[KayitDolu],">0")', CATI),
        (9, "Kritik Adet", '=COUNTIFS(tblAkis[Durum],"KRITIK",tblAkis[KayitDolu],">0")', CATI),
        (10, "Toplam Duruş Saat", "=mbk_toplamDurus", CATI),
        (11, "Duruş Maliyet", "=mbk_durusMaliyet", TL),
        (12, "Yetim Kayıt", "=mbk_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=mbk_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=mbk_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=mbk_yasKova", None),
        (16, "Günlük Duruş Ort.", "=mbk_gunlukDurusOrt", CATI),
        (17, "Tahmin Aralık", "=mbk_tahminAralik", CATI),
        (18, "Senaryo Farkı", "=mbk_senaryoKarsilastirma", CATI),
        (19, "Tornado Zirve", "=mbk_tornadoZirve", CATI),
        (20, "HHI Yoğunlaşma", "=mbk_hhi", YÜZDE),
        (21, "Kalite Skoru", "=mbk_kaliteSkor", CATI),
        (22, "P90 Yüzdelik", "=mbk_yuzdelikP90", CATI),
        (23, "Kapalı Düşüm", "=mbk_kapaliDusum", CATI),
        (24, "Haftalık Duruş", "=mbk_haftalikDurus", CATI),
        (25, "Duruş Sapma", "=mbk_durusSapma", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=mbk_yorumGunluk"),
        (8, "=mbk_yorumHaftalik"),
        (9, "=mbk_yorumSapma"),
        (10, "=mbk_yorumTornado"),
        (11, "=mbk_yorumSenaryo"),
        (12, "=mbk_yorumHhi"),
        (13, "=mbk_yorumKalite"),
        (14, "=mbk_yorumSenMotor"),
        (15, "=mbk_yorumTahmin"),
        (16, "=mbk_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Plan", 8), ("Durum Gecikme", 7), ("Durum Kritik", 5),
        ("Durum Yapıldı", 8),
    ], 29):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 28, 4, "Yaş Kova")
    h(ws, 28, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 7), ("8-30", 8), ("31-60", 6), ("60+", 7),
    ], 29):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 28, 7, "Hafta")
    h(ws, 28, 8, "Duruş Saat")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 10 + i * 2, sayi=CATI)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Duruş Saat")
    for i, (ad, t) in enumerate([
        ("Temkinli", 32), ("Baz", 38), ("İyimser", 44),
    ], 29):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=CATI)

    h(ws, 28, 13, "Hat")
    h(ws, 28, 14, "Duruş")
    for i, (ad, t) in enumerate([
        ("Hat-A", 18), ("Hat-B", 12), ("Hat-C", 8), ("Kalite", 4),
    ], 29):
        h(ws, i, 13, ad)
        h(ws, i, 14, t, sayi=CATI)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=28, max_row=32), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=29, max_row=32))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A42")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F42")

    c3 = LineChart()
    c3.title = "Haftalık Duruş"
    c3.add_data(Reference(ws, min_col=8, min_row=28, max_row=36), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=29, max_row=36))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A57")

    c4 = BarChart()
    c4.title = "Senaryo Karşılaştırma"
    c4.add_data(Reference(ws, min_col=11, min_row=28, max_row=31), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=29, max_row=31))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F57")

    c5 = PieChart()
    c5.title = "Hat Duruş Payı"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A72")

    c6 = BarChart()
    c6.title = "Kuyruk Duruş Saat"
    c6.add_data(Reference(ws, min_col=2, min_row=28, max_row=31), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=1, min_row=29, max_row=31))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F72")

    c7 = LineChart()
    c7.title = "Duruş Trend (örnek)"
    c7.add_data(Reference(ws, min_col=8, min_row=28, max_row=34), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=29, max_row=34))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A87")

    c8 = BarChart()
    c8.title = "Senaryo vs Hat"
    c8.add_data(Reference(ws, min_col=14, min_row=28, max_row=31), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=13, min_row=29, max_row=31))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F87")

    baski_hazirla(ws, "A1:N40", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "MBK-001..005 kural tablosu — 2025 / 2026 / 2027", son_kolon=12)
    kurallar_data = [
        ("MBK-001", "ISO/İç", "Planlı bakım periyodu (gün)", 30, 30, 30, "gün"),
        ("MBK-002", "ISO/Kal", "Kalibrasyon vade eşiği (gün)", 0, 0, 0, "gün"),
        ("MBK-003", "İç pol.", "Kritik gecikme eşiği (gün)", 14, 14, 14, "gün"),
        ("MBK-004", "İşyeri", "Duruş maliyet katsayısı", 1.0, 1.0, 1.0, "oran"),
        ("MBK-005", "İşyeri", "İmza/kanıt zorunluluğu (YAPILDI)", 1, 1, 1, "evet"),
    ]
    basliklar = ["kural_id", "madde", "aciklama", "y2025", "y2026", "y2027", "birim", "AktifDeger"]
    baslik_satiri(ws, 5, basliklar)
    for i, row in enumerate(kurallar_data):
        r = 6 + i
        for c, v in enumerate(row, 1):
            if c == 1:
                h(ws, r, c, v)
            elif c in (4, 5, 6) and isinstance(v, float) and v <= 2:
                _sari(ws, r, c, v, sayi=YÜZDE, baslik=row[0], mesaj="Yıl değeri")
            elif c in (4, 5, 6):
                _sari(ws, r, c, v, sayi=CATI, baslik=row[0], mesaj="Yıl değeri")
            else:
                h(ws, r, c, v)
    form_k = {
        "AktifDeger": (
            'IF(mbk_aktifYil=mbk_yilA,tblKurallar[[#This Row],[y2025]],'
            'IF(mbk_aktifYil=mbk_yilC,tblKurallar[[#This Row],[y2027]],'
            'tblKurallar[[#This Row],[y2026]]))'
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A5:H10", basliklar, formuller=form_k)
    h(ws, 12, 1, "Aktif yıl", kalin=True)
    _sari(ws, 12, 2, 2026, sayi=CATI, baslik="Aktif yıl", mesaj="2025/2026/2027")
    h(ws, 13, 1, "Kural yıl matrisi", kalin=True)
    h(ws, 13, 2, "=mbk_kuralYilMatris", kaydir=True)
    ws.merge_cells("B13:F13")
    genislik(ws, {"A": 12, "B": 12, "C": 42, "D": 10, "E": 10, "F": 10, "G": 10})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Örnek vakalar — beklenen karar ve fark kontrolü", son_kolon=10)
    basliklar = ["vaka_id", "aciklama", "durus_saat", "esik", "beklenen_karar", "fark"]
    baslik_satiri(ws, 5, basliklar)
    vakalar_data = [
        ("V01", "Plan dahilinde bakım", 1.0, 4.0, "PLANLI"),
        ("V02", "Hafif gecikme / eşik altı", 2.5, 4.0, "PLANLI"),
        ("V03", "Duruş eşik üstü", 5.0, 4.0, "GECİKMİŞ"),
        ("V04", "Kritik gecikme senaryosu", 8.0, 4.0, "KRİTİK"),
        ("V05", "Boş veri seti", 0, 4.0, "VERİ YOK"),
        ("V06", "Zincir sapmalı fiili", 4.5, 4.0, "GECİKMİŞ"),
    ]
    for i, (vid, acik, durus, esik, karar) in enumerate(vakalar_data):
        r = 6 + i
        _sari(ws, r, 1, vid, baslik="Vaka", mesaj="Vaka kimliği")
        _sari(ws, r, 2, acik, baslik="Açıklama", mesaj="Vaka açıklaması")
        _sari(ws, r, 3, durus, sayi=CATI, baslik="Duruş", mesaj="Duruş saat")
        _sari(ws, r, 4, esik, sayi=CATI, baslik="Eşik", mesaj="Eşik saat")
        _sari(ws, r, 5, karar, baslik="Beklenen", mesaj="Beklenen karar")
    form_v = {
        "fark": (
            'IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IF(AND(tblVakalar[[#This Row],[durus_saat]]=0,'
            'tblVakalar[[#This Row],[vaka_id]]="V05"),"VERİ YOK",'
            'IF(tblVakalar[[#This Row],[durus_saat]]>mbk_vakaKritikEsik,"KRİTİK",'
            'IF(tblVakalar[[#This Row],[durus_saat]]>'
            'tblVakalar[[#This Row],[esik]],"GECİKMİŞ","PLANLI"))))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:F11", basliklar, formuller=form_v)
    h(ws, 13, 1, "Vaka fark özeti", kalin=True)
    h(ws, 13, 2,
      '=COUNTIF(tblVakalar[fark],"<>"&"")&" satır"', kaydir=True)
    genislik(ws, {"A": 10, "B": 32, "C": 12, "D": 10, "E": 16, "F": 14})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Makine Bakım / Kalibrasyon / Duruş Kanıt Raporu", boyut=16, son_kolon=10)
    for i, (etiket, deger) in enumerate(rapor_baslik_satirlari(URUN_AD, SURUM), 4):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger)
    h(ws, 8, 1, "Rapor Tarihi")
    h(ws, 8, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 9, 1, "KARAR", kalin=True)
    h(ws, 9, 2, "=PANO!B4", kalin=True, boyut=14)

    ozet = [
        (11, "Açık Kuyruk", "=mbk_acikAdet", CATI),
        (12, "Toplam Duruş Saat", "=mbk_toplamDurus", CATI),
        (13, "Duruş Maliyet", "=mbk_durusMaliyet", TL),
        (14, "Yetim", "=mbk_yetimKayit", CATI),
        (15, "Geçersiz Geçiş", "=mbk_gecersizGecis", CATI),
        (16, "Zincir Kırık", "=mbk_zincirKirik", CATI),
        (17, "Kural Matris", "=mbk_kuralYilMatris", None),
        (18, "Kanıt Özeti", "=mbk_kanitRaporu", None),
        (19, "Haftalık Duruş", "=mbk_haftalikDurus", CATI),
        (20, "Kalite Skoru", "=mbk_kaliteSkor", CATI),
        (21, "Kapalı Düşüm", "=mbk_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)

    h(ws, 23, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1,
      "Duruş maliyeti = fiili duruş × saatlik maliyet × mbk_durusKatsayi. "
      "Kalibrasyon aşımı KalibrasyonGunKalan < 0 ile tespit edilir. "
      "Durum geçişleri makrosuz uyarı üretir; engellemez.",
      kaydir=True)
    ws.merge_cells("A24:F24")

    h(ws, 26, 1, "Uyarılar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1,
      "Bu çıktı karar destek belgesidir; hukuki/teknik nihai görüş yerine geçmez. "
      "Parametreler AYARLAR'da kaynak ve yürürlük tarihli tutulur.",
      kaydir=True)
    ws.merge_cells("A27:F27")

    for i, (etiket, deger) in enumerate(imza_alani(), 30):
        h(ws, i, 1, etiket)
        if deger:
            h(ws, i, 2, deger, sayi=TARİH if "Tarih" in etiket else None)
        else:
            _sari(ws, i, 2, None, baslik=etiket, mesaj="İmza/kaşe alanı")
    h(ws, 36, 1, sha_placeholder_notu(), kaydir=True, yazi=GRİ, boyut=9)
    ws.merge_cells("A36:F36")
    baski_hazirla(ws, "A1:F40", f"{URUN_AD} | Kanıt | {SURUM}")
    genislik(ws, {"A": 22, "B": 40})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Dosya sağlığı — formül ve veri kontrolleri", son_kolon=10)
    h(ws, 5, 1, "Kontrol", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    satirlar = [
        (6, "Toplam bakım", '=COUNTA(tblAkis[BakimId])'),
        (7, "Dolu kayıt", "=SUM(tblAkis[KayitDolu])"),
        (8, "Plan", '=COUNTIF(tblAkis[Durum],"PLAN")'),
        (9, "Gecikme", '=COUNTIF(tblAkis[Durum],"GECIKME")'),
        (10, "Kritik", '=COUNTIF(tblAkis[Durum],"KRITIK")'),
        (11, "Yapıldı", '=COUNTIF(tblAkis[Durum],"YAPILDI")'),
        (12, "Açık kuyruk", "=mbk_acikAdet"),
        (13, "Yetim", "=mbk_yetimKayit"),
        (14, "Geçersiz", "=mbk_gecersizGecis"),
        (15, "Zincir kırık", "=mbk_zincirKirik"),
        (16, "Veri durumu",
         '=IF(COUNTA(tblAkis[BakimId])=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=mbk_kaliteSkor"),
        (18, "Toplam duruş", "=mbk_toplamDurus"),
        (19, "Duruş maliyet", "=mbk_durusMaliyet"),
        (20, "Kapalı düşüm", "=mbk_kapaliDusum"),
    ]
    for r, ad, form in satirlar:
        h(ws, r, 1, ad)
        fmt = TL if "maliyet" in ad.lower() else (
            None if "durum" in ad.lower() else CATI)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    genislik(ws, {"A": 22, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Demo veri özeti — AKIS satırlarından türetilir", son_kolon=10)
    h(ws, 5, 1, "Alan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 6, 1, "Demo makine kartı")
    h(ws, 6, 2, DEMO_KART, sayi=CATI)
    h(ws, 7, 1, "Demo bakım satırı")
    h(ws, 7, 2, DEMO_AKIS, sayi=CATI)
    h(ws, 8, 1, "Durumlar: Plan/Gecikme/Kritik/Yapıldı — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 20000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
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

    h(ws, 3, 9, "Bakım Türü", kalin=True)
    h(ws, 6, 9, "Tur", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(BAKIM_TURLERI, 7):
        _sari(ws, i, 9, b, baslik="Tür", mesaj="Bakım türü listesi")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 1, "saat", "İş kuralı — duruş sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği (saat)"),
        ("mbk_esikDurusSaat", 4, "saat", "İç politika / MBK", "01.01.2026", "Duruş eşik saati"),
        ("mbk_kritikGecikmeGun", 14, "gün", "MBK-003", "01.01.2026", "Kritik gecikme eşiği"),
        ("mbk_kalibrasyonEsik", 0, "gün", "MBK-002", "01.01.2026", "Kalibrasyon vade eşiği"),
        ("mbk_kovaEsik1", 7, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("mbk_kovaEsik2", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("mbk_kovaEsik3", 60, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("mbk_durusKatsayi", 1.0, "oran", "MBK-004", "01.01.2026", "Duruş maliyet katsayısı"),
        ("mbk_haftalikCarpan", 0.2, "oran", "Haftalık proxy", "01.01.2026", "Haftalık duruş çarpan"),
        ("mbk_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("mbk_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("mbk_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("mbk_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("mbk_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("mbk_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("mbk_azamiDurus", 24, "saat", "İç politika — uç değer", "01.01.2026", "Azami fiili duruş"),
        ("mbk_kaliteTaban", 100, "puan", "Kalite modeli", "01.01.2026", "Kalite taban puanı"),
        ("mbk_kaliteCeza", 10, "puan", "Kalite modeli", "01.01.2026", "Yetim/geçiş/zincir cezası"),
        ("mbk_kaliteSapma", 2, "puan", "Kalite modeli", "01.01.2026", "Sapma cezası"),
        ("mbk_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("mbk_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("mbk_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("mbk_kanitRaporu_metin", "KANIT_RAPORU çıktısı", "metin", "A4 kanıt", "01.01.2026",
         "Kanıt raporu"),
        ("mbk_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("mbk_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("mbk_aktifYil", 2026, "yil", "KURALLAR", "01.01.2026", "Aktif hesap yılı"),
        ("mbk_yilA", 2025, "yil", "KURALLAR matris", "01.01.2026", "Kural yılı A"),
        ("mbk_yilC", 2027, "yil", "KURALLAR matris", "01.01.2026", "Kural yılı C"),
        ("mbk_vakaKritikEsik", 8, "saat", "VAKALAR kritik eşik", "01.01.2026",
         "Vaka kritik duruş eşiği"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.5:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    h(ws, 38, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 39, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 40 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H39:L43", dh_bas)

    h(ws, 46, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 47, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 47, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (48, "mbk_acikAdet", "=HESAP!C16"),
        (49, "mbk_toplamDurus", "=HESAP!C8"),
        (50, "mbk_durusEsikAsim", "=HESAP!C28"),
        (51, "mbk_durusMaliyet", "=HESAP!C30"),
        (52, "mbk_yetimKayit", "=HESAP!C19"),
        (53, "mbk_gecersizGecis", "=HESAP!C20"),
        (54, "mbk_zincirKirik", "=HESAP!C21"),
        (55, "mbk_kritikAdet", "=HESAP!C53"),
        (56, "mbk_yasKova", "=HESAP!C49"),
        (57, "mbk_gunlukDurusOrt", "=HESAP!C27"),
        (58, "mbk_durusSapma", "=HESAP!C33"),
        (59, "mbk_tahminAralik", "=HESAP!C43"),
        (60, "mbk_senaryoKarsilastirma", "=HESAP!C37"),
        (61, "mbk_tornadoZirve", "=HESAP!C38"),
        (62, "mbk_hhi", "=HESAP!C39"),
        (63, "mbk_kaliteSkor", "=HESAP!C40"),
        (64, "mbk_yuzdelikP90", "=HESAP!C44"),
        (65, "mbk_senaryoMotor", "=HESAP!C50"),
        (66, "mbk_kapaliDusum", "=HESAP!C17"),
        (67, "mbk_kuyrukOzeti", "=HESAP!C48"),
        (68, "mbk_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C16,"Yetim",HESAP!C19,'
         '"Zincir",HESAP!C21,"Duruş",HESAP!C8,"Kritik",HESAP!C53)'),
        (69, "mbk_sonrakiKimlik",
         sonraki_kimlik_formul("MAK", "tblMakineler[MakineId]").replace(
             'MAX(tblMakineler[MakineId])',
             'COUNTA(tblMakineler[MakineId])')),
        (70, "mbk_haftalikDurus", "=HESAP!C32"),
        (71, "mbk_gecikmeAdet", "=HESAP!C52"),
        (72, "mbk_kuralYilMatris",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"MBK-001..005","2025-2027",mbk_aktifYil)'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 74, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (75, "mbk_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük duruş ort",TEXT(mbk_gunlukDurusOrt,"0.0"),"saat")'),
        (76, "mbk_yorumHaftalik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Haftalık duruş proxy",TEXT(mbk_haftalikDurus,"0.0"),"saat")'),
        (77, "mbk_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Duruş sapma",TEXT(mbk_durusSapma,"0.00"),"saat")'),
        (78, "mbk_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(mbk_tornadoZirve,"0.00"),"saat")'),
        (79, "mbk_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(mbk_senaryoKarsilastirma,"0.00"))'),
        (80, "mbk_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(mbk_hhi,"0.0%"))'),
        (81, "mbk_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",mbk_kaliteSkor)'),
        (82, "mbk_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",mbk_senaryoMotor)'),
        (83, "mbk_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(mbk_tahminAralik,"0.00"),"saat")'),
        (84, "mbk_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 duruş",TEXT(mbk_yuzdelikP90,"0.00"),"saat")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. MAKINELER sayfasına makine kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına plan tarih, fiili duruş, kalibrasyon kalan ve durum satırını yazın.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — görünür uyarı üretir (makrosuz kural).",
        "5. Kapalı durum (Yapıldı) açık kuyruklardan düşer.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma plan tarihine göredir.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. KANIT_RAPORU sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 36):
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
    ws = wb["KANIT_RAPORU"]
    for r, ad in [(5, "Sürüm"), (30, "İmza1"), (31, "İmza2"), (32, "İmza3")]:
        dogrulama(ws, "textLength", "0", f"B{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KILAVUZ"]
    for r in range(5, 14):
        dogrulama(ws, "textLength", "0", f"A{r}",
                  baslik="Kılavuz", mesaj="Bilgi satırı",
                  hata_baslik="Metin", hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KURALLAR"]
    for r in range(6, 11):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Kural", mesaj="MBK-###",
                  hata_baslik="Kural", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="20")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"PLANLI"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"GECİKMİŞ"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 26):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["AKIS"]
    for col, val, fill, font in [
        ("R", '"YETİM"', kirmizi, kf),
        ("S", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("T", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("T", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("GECIKME", sari), ("KRITIK", kirmizi), ("YAPILDI", yesil)]:
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
    for r in range(6, 54):
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

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"PLANLI"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"GECİKMİŞ"'], fill=sari))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(11, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["MAKINELER"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"E{ILK}:E{SON}",
        CellIsRule(operator="greaterThan", formula=["1000"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "mbk_esikDurusSaat", "mbk_kritikGecikmeGun",
        "mbk_kalibrasyonEsik", "mbk_kovaEsik1", "mbk_kovaEsik2", "mbk_kovaEsik3",
        "mbk_durusKatsayi", "mbk_haftalikCarpan",
        "mbk_senaryoTemkinli", "mbk_senaryoBaz", "mbk_senaryoIyimser",
        "mbk_tornadoOran", "mbk_yuzdelikOran", "mbk_olcekHedef", "mbk_azamiDurus",
        "mbk_kaliteTaban", "mbk_kaliteCeza", "mbk_kaliteSapma",
        "mbk_durumHaritasi", "mbk_kuyrukOzeti_metin", "mbk_kapaliDusum_metin",
        "mbk_kanitRaporu_metin", "mbk_dosyaSurumu", "mbk_firmaUnvan", "mbk_aktifYil",
        "mbk_yilA", "mbk_yilC", "mbk_vakaKritikEsik",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "mbk_acikAdet": 48, "mbk_toplamDurus": 49, "mbk_durusEsikAsim": 50,
        "mbk_durusMaliyet": 51, "mbk_yetimKayit": 52, "mbk_gecersizGecis": 53,
        "mbk_zincirKirik": 54, "mbk_kritikAdet": 55, "mbk_yasKova": 56,
        "mbk_gunlukDurusOrt": 57, "mbk_durusSapma": 58, "mbk_tahminAralik": 59,
        "mbk_senaryoKarsilastirma": 60, "mbk_tornadoZirve": 61, "mbk_hhi": 62,
        "mbk_kaliteSkor": 63, "mbk_yuzdelikP90": 64, "mbk_senaryoMotor": 65,
        "mbk_kapaliDusum": 66, "mbk_kuyrukOzeti": 67, "mbk_kanitRaporu": 68,
        "mbk_sonrakiKimlik": 69, "mbk_haftalikDurus": 70, "mbk_gecikmeAdet": 71,
        "mbk_kuralYilMatris": 72,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "mbk_yorumGunluk": 75, "mbk_yorumHaftalik": 76, "mbk_yorumSapma": 77,
        "mbk_yorumTornado": 78, "mbk_yorumSenaryo": 79, "mbk_yorumHhi": 80,
        "mbk_yorumKalite": 81, "mbk_yorumSenMotor": 82, "mbk_yorumTahmin": 83,
        "mbk_yorumP90": 84,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$10")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$G$7:$G$8")
    ad_ekle(wb, "ListeBakimTuru", "LISTELER!$I$7:$I$10")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (makineler, "MAKINELER"),
        (akis, "AKIS"),
        (kuyruklar, "KUYRUKLAR"),
        (hesap, "HESAP"),
        (pano, "PANO"),
        (kurallar, "KURALLAR"),
        (vakalar, "VAKALAR"),
        (kanit_raporu, "KANIT_RAPORU"),
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
    dosya_ad = "MakineBakimKalibrasyonDurusMaliyeti.xlsx"
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
