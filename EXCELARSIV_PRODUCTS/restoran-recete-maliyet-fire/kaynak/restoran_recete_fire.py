#!/usr/bin/env python3
"""Restoran Reçete Maliyet + Fire Hesaplama — A7 reçete motoru (manda v6).

Tek restoran odak reçetesinden fire kolonlu porsiyon maliyeti üretir.
Menü mühendisliği matrisi YOK. ortak.recete_motoru + rrf_ önek.
R01 %100 bileşim · R02 birim dönüşüm · R03 fire · R04 versiyon · R05 kırılım.
Karar: VERİ YOK / UYGUN / DİKKAT / KRİTİK.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from ortak.recete_motoru import (  # noqa: E402
    bilesim_yuzde_alarm,
    fire_dahil_miktar,
    varyans_koprusu,
)

from excel_uretim.ortak import (  # noqa: E402
    CATI,
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

URUN_AD = "Restoran Reçete Maliyet + Fire Hesaplama"
SURUM = "1.0.0"
RENK = "9A3412"
KAPASITE = 1000
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 14)


# Demo: odak menü kalemi (Adana Kebap porsiyon) bileşim %100
HAMMADDELER = [
    ("HM-01", "Dana Kiyma", "kg", "kg", 420.0, 0.05, "Et Tedarik A.S."),
    ("HM-02", "Kuzu Yag", "kg", "kg", 280.0, 0.03, "Et Tedarik A.S."),
    ("HM-03", "Sogan", "kg", "kg", 28.0, 0.08, "Sebze Hali"),
    ("HM-04", "Biber", "kg", "kg", 55.0, 0.06, "Sebze Hali"),
    ("HM-05", "Baharat Karisim", "kg", "kg", 180.0, 0.01, "Baharatci"),
    ("HM-06", "Zeytinyagi", "lt", "ml", 220.0, 0.02, "Yag Depo"),
    ("HM-07", "Lavas", "adet", "adet", 8.0, 0.02, "Firin"),
]

RECETE_SATIR = [
    ("RC-01", "HM-01", 0.55, 0.05, "V2", "Ana protein"),
    ("RC-02", "HM-02", 0.12, 0.03, "V2", "Yag"),
    ("RC-03", "HM-03", 0.10, 0.08, "V2", "Sogan"),
    ("RC-04", "HM-04", 0.08, 0.06, "V2", "Biber"),
    ("RC-05", "HM-05", 0.03, 0.01, "V2", "Baharat"),
    ("RC-06", "HM-06", 0.05, 0.02, "V2", "Yag"),
    ("RC-07", "HM-07", 0.07, 0.02, "V2", "Lavas"),
    ("RC-08", "HM-01", 0.50, 0.06, "V1", "Eski daha az et"),
    ("RC-09", "HM-02", 0.15, 0.04, "V1", "Eski yag"),
    ("RC-10", "HM-03", 0.12, 0.08, "V1", "Eski sogan"),
    ("RC-11", "HM-04", 0.08, 0.06, "V1", "Eski biber"),
    ("RC-12", "HM-05", 0.04, 0.01, "V1", "Eski baharat"),
    ("RC-13", "HM-06", 0.05, 0.02, "V1", "Eski yag"),
    ("RC-14", "HM-07", 0.06, 0.02, "V1", "Eski lavas"),
]

DONUSUM = [
    ("kg_kg", "kg", "kg", 1.0, "Kimlik"),
    ("lt_ml", "lt", "ml", 1000.0, "Metroloji"),
    ("adet_adet", "adet", "adet", 1.0, "Kimlik"),
    ("cuval_kg", "cuval", "kg", 25.0, "Tedarik sozlesmesi 25 kg"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Porsiyon maliyeti ve menü kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı maliyet üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _kim(ws, hucre, baslik, mesaj):
    yorum_ekle(
        ws, hucre,
        f"Tanım: {baslik} | Neden önemli: Menü reçetesi ve porsiyon maliyetini etkiler | "
        f"Doğru kullanım: {mesaj} | Örnek: demo satırına bakın | "
        f"Yanlışsa: porsiyon maliyeti bozulur | Kaynak: Mutfak / AYARLAR",
    )


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    h(ws, 3, 1, URUN_AD, kalin=True, boyut=18, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      "Tek restoran odak reçetesini kurun; fire kolonlu porsiyon maliyetini hesaplayın; "
      "bileşim %100 ve versiyon varyansıyla fiyat önerisi alın.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:N4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Reçete bileşim toplamı %100 kontrolü — sapmada PANO kırmızı alarm (R01)",
        "Fire kolonlu porsiyon maliyeti — rrf_fireDahilMaliyet (R03)",
        "AYARLAR birim dönüşüm tablosu — gömülü katsayı yok (R02)",
        "Versiyon varyans köprüsü V1↔V2 — rrf_varyansKopru (R04)",
        "Makrosuz; karar VERİ YOK / UYGUN / DİKKAT / KRİTİK — menü matrisi yok",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Şef — reçete bileşimini ve fire'yi kilitler",
        "Restoran sahibi — porsiyon maliyeti ve fiyat önerisini görür",
        "Mali işler — fire maliyet kanıtını dosyalar",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI'ye ürün ve satış fiyatını girin. 2) HAMMADDE fiyat/fire güncelleyin. "
      "3) RECETE bileşimini kilitleyin (%100). 4) PANO ve KARAR'dan sonucu görün.",
      kaydir=True)
    ws.merge_cells("A19:N19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Kod: RRF-PRO | Lisans: Tek kullanıcı | ExcelArşiv",
      yazi=GRİ, boyut=9)
    genislik(ws, {"A": 72})




def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", RENK, URUN_AD, son_kolon=18)
    h(ws, 3, 1, "İşletme ve odak menü kalemi girdileri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:G3")
    h(ws, 4, 1, "Sarı alanlar manuel giriş — liste ve sayı doğrulamalı", yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A4:G4")

    alanlar = [
        (6, "SirketAdi", "Örnek Restoran", None, "İşletme adı", "Metin girin"),
        (7, "Donem", "2026-08", None, "Dönem", "YYYY-AA"),
        (8, "MenuKalemi", "Adana Kebap", None, "Odak menü kalemi", "Menü adı"),
        (9, "GunlukSatisAdet", 80, CATI, "Günlük satış adedi", "Pozitif tam sayı"),
        (10, "PorsiyonGram", 350.0, "0.0", "Porsiyon gramajı (g)", "0 üstü sayı"),
        (11, "SatisFiyat", 320.0, TL, "Satış fiyatı (TL/porsiyon)", "0 üstü TL"),
        (12, "Populerlik", 0.28, YÜZDE, "Popülerlik (satış payı)", "0-1 arası"),
        (13, "IscilikTlPorsiyon", 18.0, TL, "İşçilik (TL/porsiyon)", "0 üstü TL"),
        (14, "AmbalajTlPorsiyon", 4.0, TL, "Ambalaj / servis (TL/porsiyon)", "0 üstü TL"),
        (15, "GenelGiderTlPorsiyon", 12.0, TL, "Genel gider payı (TL/porsiyon)", "0 üstü TL"),
        (16, "HedefMarj", 0.65, YÜZDE, "Hedef brüt marj", "0-1 arası"),
        (17, "AktifVersiyon", "V2", None, "Aktif reçete versiyonu", "V1 veya V2"),
    ]
    h(ws, 5, 1, "Alan")
    h(ws, 5, 2, "Deger", kalin=True)
    h(ws, 5, 3, "Birim")
    for r, anahtar, deger, sayi, baslik, mesaj in alanlar:
        h(ws, r, 1, anahtar)
        _sari(ws, r, 2, deger, sayi=sayi, baslik=baslik, mesaj=mesaj)
        h(ws, r, 3, {
            "GunlukSatisAdet": "adet", "PorsiyonGram": "g", "SatisFiyat": "TL/porsiyon",
            "Populerlik": "oran", "IscilikTlPorsiyon": "TL/porsiyon",
            "AmbalajTlPorsiyon": "TL/porsiyon", "GenelGiderTlPorsiyon": "TL/porsiyon",
            "HedefMarj": "oran", "AktifVersiyon": "kod", "MenuKalemi": "metin",
            "Donem": "donem", "SirketAdi": "metin",
        }.get(anahtar, ""))

    baslik_satiri(ws, 20, [
        "AlanAnahtari", "Deger", "Birim", "Zorunlu", "Kaynak", "Not", "DoluBayrak",
    ])
    demo_satir = [
        ("MenuKalemi", "Adana Kebap", "metin", "EVET", "Mutfak", "Demo"),
        ("GunlukSatisAdet", 80, "adet", "EVET", "Kasa", "Demo"),
        ("PorsiyonGram", 350, "g", "EVET", "Reçete", "Demo"),
        ("SatisFiyat", 320, "TL/porsiyon", "EVET", "Menü", "Demo"),
        ("Populerlik", 0.28, "oran", "EVET", "Satış", "Demo"),
        ("IscilikTlPorsiyon", 18, "TL/porsiyon", "EVET", "Bordro", "Demo"),
        ("AmbalajTlPorsiyon", 4, "TL/porsiyon", "HAYIR", "Depo", "Demo"),
        ("GenelGiderTlPorsiyon", 12, "TL/porsiyon", "EVET", "Muhasebe", "Demo"),
        ("HedefMarj", 0.65, "oran", "EVET", "CFO", "Demo"),
        ("AktifVersiyon", "V2", "kod", "EVET", "Reçete", "Demo"),
    ]
    for i, (a, d, b, z, k, n) in enumerate(demo_satir):
        r = 21 + i
        _sari(ws, r, 1, a, baslik="Alan anahtarı", mesaj="Girdi anahtarı")
        if isinstance(d, float) and a in ("HedefMarj", "Populerlik"):
            _sari(ws, r, 2, d, sayi=YÜZDE, baslik=a, mesaj="Değer")
        elif isinstance(d, (int, float)) and a in (
            "SatisFiyat", "IscilikTlPorsiyon", "AmbalajTlPorsiyon", "GenelGiderTlPorsiyon"
        ):
            _sari(ws, r, 2, d, sayi=TL, baslik=a, mesaj="Değer")
        elif isinstance(d, (int, float)):
            _sari(ws, r, 2, d, sayi=CATI if isinstance(d, int) else "0.0", baslik=a, mesaj="Değer")
        else:
            _sari(ws, r, 2, d, baslik=a, mesaj="Değer")
        _sari(ws, r, 3, b, baslik="Birim", mesaj="Ölçü birimi")
        _sari(ws, r, 4, z, baslik="Zorunlu", mesaj="EVET/HAYIR")
        _sari(ws, r, 5, k, baslik="Kaynak", mesaj="Veri kaynağı")
        _sari(ws, r, 6, n, baslik="Not", mesaj="Açıklama")
        h(ws, r, 7, None)
    for r in range(21 + len(demo_satir), SON + 1):
        for c in range(1, 7):
            _sari(ws, r, c, None, baslik="Girdi alanı", mesaj="Boş satır — doldurulabilir")
        h(ws, r, 7, None)

    form = {
        "DoluBayrak": (
            'IF(tblGirdi[[#This Row],[AlanAnahtari]]="","",'
            'IF(OR(tblGirdi[[#This Row],[Deger]]="",tblGirdi[[#This Row],[Birim]]=""),'
            '"EKSIK","DOLU"))'
        ),
    }
    tablo_ekle(ws, "tblGirdi", f"A20:G{SON}", [
        "AlanAnahtari", "Deger", "Birim", "Zorunlu", "Kaynak", "Not", "DoluBayrak",
    ], form)

    dogrulama(ws, "list", '"EVET,HAYIR"', f"D21:D{SON}",
              baslik="Zorunlu", mesaj="EVET veya HAYIR seçin",
              hata_baslik="Geçersiz", hata_mesaj="Yalnız EVET/HAYIR")
    dogrulama(ws, "list", "=ListeVersiyon", "B17",
              baslik="Versiyon", mesaj="V1 veya V2",
              hata_baslik="Geçersiz", hata_mesaj="V1/V2")
    dogrulama(ws, "decimal", "0", "B9:B16",
              baslik="Sayı", mesaj="0 veya üzeri",
              hata_baslik="Negatif olamaz", hata_mesaj="0+ girin",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "=ListeEvetHayir", f"D21:D{SON}",
              baslik="Zorunlu bayrak", mesaj="EVET/HAYIR",
              hata_baslik="Uyarı", hata_mesaj="Değeri kontrol edin")

    sabitle(ws, "A21")
    genislik(ws, {"A": 55, "B": 22, "C": 14, "D": 10, "E": 14, "F": 18, "G": 12})


def hammadde(ws):
    sayfa_hazirla(ws, "HAMMADDE", RENK, URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Mutfak hammadde kartı — alım birimi → kullanım birimi dönüşümü AYARLAR'dan",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A3:L3")
    sutunlar = [
        "KalemId", "KalemAd", "AlimBirim", "KullanimBirim", "AlimFiyat_tl",
        "Fire", "Tedarikci", "Guncelleme", "Katsayi", "NetKgFiyat_tl", "Not",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "Katsayi": (
            'IF(tblHammadde[[#This Row],[KalemId]]="","",'
            'IFERROR(INDEX(tblDonusum[katsayi],MATCH('
            'IF(tblHammadde[[#This Row],[AlimBirim]]="lt","lt_ml",'
            'IF(tblHammadde[[#This Row],[AlimBirim]]="adet","adet_adet",'
            'IF(tblHammadde[[#This Row],[AlimBirim]]="cuval","cuval_kg","kg_kg"))),'
            'tblDonusum[donusum_id],0)),0))'
        ),
        "NetKgFiyat_tl": (
            'IF(tblHammadde[[#This Row],[KalemId]]="","",'
            'IFERROR(tblHammadde[[#This Row],[AlimFiyat_tl]]/'
            'MAX(tblHammadde[[#This Row],[Katsayi]],rrf_sifirKoruma)'
            '*(1+tblHammadde[[#This Row],[Fire]]),0))'
        ),
    }
    for i, row in enumerate(HAMMADDELER):
        r = ILK + i
        kid, ad, alim, kul, fiyat, fire, ted = row
        _sari(ws, r, 1, kid, baslik="Kalem kimliği", mesaj="HM-xx")
        _sari(ws, r, 2, ad, baslik="Kalem adı", mesaj="Malzeme adı")
        _sari(ws, r, 3, alim, baslik="Alım birimi", mesaj="kg/lt/adet")
        _sari(ws, r, 4, kul, baslik="Kullanım birimi", mesaj="kg")
        _sari(ws, r, 5, fiyat, sayi=TL, baslik="Alım fiyatı", mesaj="Alım birimi fiyatı TL")
        _sari(ws, r, 6, fire, sayi=YÜZDE, baslik="Fire", mesaj="0-1 fire oranı")
        _sari(ws, r, 7, ted, baslik="Tedarikçi", mesaj="Tedarikçi adı")
        _sari(ws, r, 8, RAPOR_TARIHI, sayi=TARİH, baslik="Güncelleme", mesaj="Fiyat tarihi")
        h(ws, r, 9, None)  # formül
        h(ws, r, 10, None, sayi=TL)
        _sari(ws, r, 11, "Demo", baslik="Not", mesaj="Serbest not")
    for r in range(ILK + len(HAMMADDELER), SON + 1):
        for c in range(1, 12):
            _sari(ws, r, c, None, baslik="Hammadde alanı", mesaj="Yeni kalem girilebilir",
                  sayi=(TL if c in (5, 10) else YÜZDE if c == 6 else TARİH if c == 8 else None))

    tablo_ekle(ws, "tblHammadde", f"A{HDR}:K{SON}", sutunlar, form)
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}",
              baslik="Alım fiyatı", mesaj="0 veya üzeri TL",
              hata_baslik="Negatif olamaz", hata_mesaj="0+ girin",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Fire", mesaj="0-1 arası",
              hata_baslik="Geçersiz fire", hata_mesaj="0-1 girin",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "=ListeBirim", f"C{ILK}:D{SON}",
              baslik="Birim", mesaj="Listeden seçin",
              hata_baslik="Geçersiz", hata_mesaj="kg/lt/adet")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["G"].width = 16


def recete(ws):
    sayfa_hazirla(ws, "RECETE", RENK, URUN_AD, son_kolon=18)
    h(ws, 3, 1, "Reçete bileşim tablosu — oran toplamı %100 (aktif versiyon)", kaydir=True, yazi=GRİ)
    ws.merge_cells("A3:J3")

    h(ws, 4, 1, "Aktif versiyon bileşim toplamı %")
    h(ws, 4, 2,
      '=IFERROR(SUMIF(tblRecete[Versiyon],GIRDI!B17,tblRecete[Oran])*100,0)',
      sayi="0.00", kalin=True)
    _kim(ws, "B4", "Bileşim toplamı", "Aktif versiyon oranları ×100")
    h(ws, 4, 3, "Bileşim alarmı")
    h(ws, 4, 4,
      bilesim_yuzde_alarm("B4", "rrf_bilesimTolerans"),
      kalin=True)
    _kim(ws, "D4", "Bileşim %100 alarmı", "Tolerans dışı sapmada ALARM")
    h(ws, 4, 5, "Fire brüt kg")
    h(ws, 4, 6, fire_dahil_miktar("GIRDI!B10/1000", "C6"), sayi="0.000")
    _kim(ws, "F4", "Fire dahil miktar", "ortak.recete_motoru fire_dahil_miktar")

    sutunlar = [
        "SatirId", "KalemId", "Oran", "Fire", "Versiyon", "MiktarKg", "MaliyetPay_tl", "Not",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "MiktarKg": (
            'IF(tblRecete[[#This Row],[SatirId]]="","",'
            'IFERROR((GIRDI!B10/1000)*tblRecete[[#This Row],[Oran]]'
            '/(1-tblRecete[[#This Row],[Fire]]),0))'
        ),
        "MaliyetPay_tl": (
            'IF(tblRecete[[#This Row],[SatirId]]="","",'
            'IFERROR(tblRecete[[#This Row],[MiktarKg]]*'
            'IFERROR(INDEX(tblHammadde[NetKgFiyat_tl],MATCH('
            'tblRecete[[#This Row],[KalemId]],tblHammadde[KalemId],0)),0),0))'
        ),
    }
    for i, row in enumerate(RECETE_SATIR):
        r = ILK + i
        sid, kid, oran, fire, ver, notu = row
        _sari(ws, r, 1, sid, baslik="Satır kimliği", mesaj="RC-xx")
        _sari(ws, r, 2, kid, baslik="Kalem kimliği", mesaj="HM-xx")
        _sari(ws, r, 3, oran, sayi=YÜZDE, baslik="Oran", mesaj="Bileşim oranı 0-1")
        _sari(ws, r, 4, fire, sayi=YÜZDE, baslik="Fire", mesaj="Satır fire oranı")
        _sari(ws, r, 5, ver, baslik="Versiyon", mesaj="V1 veya V2")
        h(ws, r, 6, None, sayi="0.00")
        h(ws, r, 7, None, sayi=TL)
        _sari(ws, r, 8, notu, baslik="Not", mesaj="Açıklama")
    for r in range(ILK + len(RECETE_SATIR), SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, baslik="Reçete alanı", mesaj="Yeni satır",
                  sayi=(YÜZDE if c in (3, 4) else "0.00" if c == 6 else TL if c == 7 else None))

    tablo_ekle(ws, "tblRecete", f"A{HDR}:H{SON}", sutunlar, form)
    dogrulama(ws, "list", "=ListeVersiyon", f"E{ILK}:E{SON}",
              baslik="Versiyon", mesaj="V1/V2",
              hata_baslik="Geçersiz", hata_mesaj="V1 veya V2")
    dogrulama(ws, "decimal", "0", f"C{ILK}:D{SON}",
              baslik="Oran/Fire", mesaj="0-1",
              hata_baslik="Geçersiz", hata_mesaj="0+",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 9)})
    ws.column_dimensions["A"].width = 36


def maliyet(ws):
    sayfa_hazirla(ws, "MALIYET", RENK, URUN_AD, son_kolon=16)
    h(ws, 3, 1, "Maliyet motoru — hammadde + işçilik + ambalaj + genel gider → porsiyon maliyeti",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A3:J3")

    adimlar = [
        (6, "Dolu girdi satırı",
         '=IFERROR(COUNTA(tblGirdi[AlanAnahtari]),0)', CATI, "Veri varlığı"),
        (7, "Aktif reçete satırı",
         '=IFERROR(COUNTIF(tblRecete[Versiyon],GIRDI!B17),0)', CATI, "Versiyon satır sayısı"),
        (8, "Bileşim toplam %",
         '=IFERROR(SUMIF(tblRecete[Versiyon],GIRDI!B17,tblRecete[Oran])*100,0)', "0.00", "R01"),
        (9, "Bileşim alarmı",
         bilesim_yuzde_alarm("B8", "rrf_bilesimTolerans"), None, "R01 alarm"),
        (10, "Hammadde maliyet TL/porsiyon",
         '=IFERROR(SUMIF(tblRecete[Versiyon],GIRDI!B17,tblRecete[MaliyetPay_tl]),0)', TL, "Fire dahil"),
        (11, "Günlük hammadde TL",
         '=IFERROR(B10*GIRDI!B9,0)', TL, "× satış adedi"),
        (12, "İşçilik TL/porsiyon",
         '=IFERROR(GIRDI!B13,0)', TL, "Mutfak işçilik"),
        (13, "Ambalaj TL/porsiyon",
         '=IFERROR(GIRDI!B14,0)', TL, "Servis/ambalaj"),
        (14, "Genel gider TL/porsiyon",
         '=IFERROR(GIRDI!B15,0)', TL, "Genel pay"),
        (15, "Toplam maliyet TL/porsiyon",
         '=IFERROR(B10+B12+B13+B14,0)', TL, "Kırılım toplamı"),
        (16, "Hammadde birim pay TL",
         '=IFERROR(B10,0)', TL, "Hammadde"),
        (17, "İşçilik birim pay TL",
         '=IFERROR(B12,0)', TL, "İşçilik"),
        (18, "Ambalaj birim pay TL",
         '=IFERROR(B13,0)', TL, "Ambalaj"),
        (19, "Genel gider birim pay TL",
         '=IFERROR(B14,0)', TL, "Genel"),
        (20, "Porsiyon maliyeti TL",
         '=IFERROR(B16+B17+B18+B19,0)', TL, "R05 birim"),
        (21, "Kırılım kontrolü",
         '=IF(ABS((rrf_kirilimHammadde+rrf_kirilimIscilik+rrf_kirilimAmbalaj+rrf_kirilimGenel)'
         '-B20)>rrf_tolerans,"KIRILIM UYUSMUYOR","OK")', None, "R05"),
        (22, "Günlük satış adedi",
         '=IFERROR(GIRDI!B9,0)', CATI, "Adet"),
        (23, "Marj TL/porsiyon",
         '=IFERROR(GIRDI!B11-B20,0)', TL, "Satış-maliyet"),
        (24, "Marj oranı",
         '=IFERROR(B23/MAX(GIRDI!B11,rrf_sifirKoruma),0)', YÜZDE, "Marj/fiyat"),
        (25, "Önerilen fiyat TL",
         '=IFERROR(B20/(1-GIRDI!B16),0)', TL, "Hedef marjlı fiyat"),
        (26, "Gida maliyeti orani",
         '=IFERROR(B20/MAX(GIRDI!B11,rrf_sifirKoruma),0)', YÜZDE, "Maliyet/fiyat"),
        (27, "Reçete toplam kg/porsiyon",
         '=IFERROR(SUMIF(tblRecete[Versiyon],GIRDI!B17,tblRecete[MiktarKg]),0)', "0.00", "Fire dahil kg"),
        (28, "Fire dahil maliyet bayrağı",
         '=IFERROR(IF(COUNTIF(tblRecete[Fire],">=0")>0,"FIRE_VAR","FIRE_YOK"),"FIRE_YOK")', None, "R03"),
        (29, "Karar kodu",
         '=IF(B6=0,"VERI_YOK",'
         'IF(OR(B9="ALARM: bileşim %100 dışı",B26>rrf_kritikOran),"KRITIK",'
         'IF(B26>rrf_dikkatOran,"DIKKAT","UYGUN")))', None, "Kapı kodu"),
        (30, "Karar metni",
         '=IF(B29="VERI_YOK","VERİ YOK",'
         'IF(B29="KRITIK","KRİTİK",'
         'IF(B29="DIKKAT","DİKKAT","UYGUN")))', None, "Görünen karar"),
        (31, "Karar gerekçe",
         '=_xlfn.TEXTJOIN(" | ",TRUE,'
         '"Bileşim",B8,"%","Porsiyon",TEXT(B20,"0.00"),'
         '"Marj",TEXT(B24,"0.0%"),"Alarm",B9)', None, "Gerekçe"),
        (32, "Kalite skoru",
         '=IFERROR(MAX(0,MIN(100,rrf_kaliteTaban'
         '-IF(B9<>"OK",rrf_kaliteCeza,0)'
         '-IF(B21<>"OK",rrf_kaliteCeza,0)'
         '-IF(B6<rrf_girisBeklenen,rrf_kaliteSapma,0))),0)', CATI, "0-100"),
        (33, "HHI yoğunlaşma",
         '=IFERROR(SUMPRODUCT((tblRecete[Oran])^2'
         '*(tblRecete[Versiyon]=GIRDI!B17)*1),0)', "0.000", "O6"),
        (34, "Senaryo baz maliyet",
         '=IFERROR(B20,0)', TL, "Baz"),
        (35, "Senaryo iyimser",
         '=IFERROR(B20*(1-rrf_senaryoIyi),0)', TL, "İyimser"),
        (36, "Senaryo kötümser",
         '=IFERROR(B20*(1+rrf_senaryoKotu),0)', TL, "Kötümser"),
        (37, "Tornado et fiyatı etkisi",
         '=IFERROR(B20*rrf_tornadoEt,0)', TL, "Et fiyat"),
        (38, "Tornado yag etkisi",
         '=IFERROR(B20*rrf_tornadoYag,0)', TL, "Yağ fiyat"),
        (39, "Tornado popülerlik etkisi",
         '=IFERROR(B20*rrf_tornadoPop,0)', TL, "Popülerlik"),
        (40, "Tahmin alt",
         '=IFERROR(B20*rrf_tahminAlt,0)', TL, "I1 alt"),
        (41, "Tahmin üst",
         '=IFERROR(B20*rrf_tahminUst,0)', TL, "I1 üst"),
        (42, "Yüzdelik P90 ham",
         '=IFERROR(_xlfn.PERCENTILE.INC(SENARYO!C6:C8,rrf_yuzdelikOran),0)', TL, "I2"),
        (43, "Varyans toplam",
         '=IFERROR(VARYANS!B12,0)', TL, "R04 köprü"),
        (44, "Fire dahil maliyet",
         '=IFERROR(B10,0)', TL, "R03 iz"),
        (45, "Öneri aksiyon 1",
         '=IF(B29="VERI_YOK","Önce GIRDI doldurun",'
         'IF(B29="KRITIK","Et/yağ fiyatını ve bileşim %100 kontrol edin",'
         'IF(B29="DIKKAT","Fire ve protein kalemini gözden geçirin","Reçeteyi koruyun")))',
         None, "Aksiyon"),
        (46, "Öneri aksiyon 2",
         '=IF(B24<GIRDI!B16,"Hedef marj için satış fiyatını veya maliyeti ayarlayın","Marj hedefte")',
         None, "Aksiyon"),
        (47, "Öneri aksiyon 3",
         '=IF(B33>rrf_hhiEsik,"Tek kaleme bağımlılık yüksek — çeşitlendirin","Yoğunlaşma kabul")',
         None, "Aksiyon"),
        (48, "Porsiyon hammadde", '=IFERROR(B10,0)', TL, "adet"),
        (49, "Günlük toplam maliyet", '=IFERROR(B15*GIRDI!B9,0)', TL, "gün"),
        (50, "Birim maliyet çapraz", '=IFERROR(B15,0)', TL, "kontrol"),
        (51, "Sapma birim vs çapraz", '=IFERROR(ABS(B20-B50),0)', TL, "D12"),
        (52, "Menü kalemi", '=IFERROR(GIRDI!B8,"")', None, "menü"),
        (53, "Aktif versiyon kod", '=IFERROR(GIRDI!B17,"")', None, "versiyon"),
        (54, "Hammadde pay oranı", '=IFERROR(B10/MAX(B15,rrf_sifirKoruma),0)', YÜZDE, "pay"),
        (55, "İşçilik pay oranı", '=IFERROR(B12/MAX(B15,rrf_sifirKoruma),0)', YÜZDE, "pay"),
        (56, "Ambalaj pay oranı", '=IFERROR(B13/MAX(B15,rrf_sifirKoruma),0)', YÜZDE, "pay"),
        (57, "Genel pay oranı", '=IFERROR(B14/MAX(B15,rrf_sifirKoruma),0)', YÜZDE, "pay"),
        (58, "Kritik kalem uyarısı",
         '=IF(B54>rrf_kritikHammaddePay,"KRİTİK KALEM: hammadde payı yüksek","OK")', None, "uyarı"),
        (59, "FORECAST satış eğilim",
         '=IFERROR(FORECAST(rrf_fcHedef,AYARLAR!$B$45:$B$47,AYARLAR!$B$48:$B$50),0)',
         "0.0", "tahmin yardımcı"),
        (60, "TREND yardımcı",
         '=IFERROR(INDEX(TREND(AYARLAR!$B$51:$B$53,AYARLAR!$B$48:$B$50),1),0)',
         TL, "fiyat trend"),
        (61, "Fire tutari TL",
         '=IFERROR(B44,0)', TL, "rrf_fireDahilMaliyet ayna"),
    ]

    h(ws, 5, 1, "Adım", kalin=True)
    h(ws, 5, 2, "Deger", kalin=True)
    h(ws, 5, 3, "Aciklama", kalin=True)
    for r, ad, formul, sayi, acik in adimlar:
        h(ws, r, 1, ad)
        h(ws, r, 2, formul, sayi=sayi, kalin=(r in (20, 30)))
        h(ws, r, 3, acik, yazi=GRİ, boyut=9)

    # Grafik veri bloğu — D09b için sabit demo sayılar (formül değil)
    h(ws, 62, 1, "KirilimEtiket")
    h(ws, 62, 2, "KirilimDeger")
    for i, (etiket, deger) in enumerate([
        ("Hammadde", 78.0),
        ("Iscilik", 18.0),
        ("Ambalaj", 4.0),
        ("Genel gider", 12.0),
    ], 63):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger, sayi=TL)

    h(ws, 68, 1, "SenaryoEtiket")
    h(ws, 68, 2, "SenaryoDeger")
    for i, (etiket, deger) in enumerate([
        ("Iyimser", 95.0),
        ("Baz", 112.0),
        ("Kotumser", 128.0),
    ], 69):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger, sayi=TL)

    # canlı ayna (grafik dışı)
    h(ws, 73, 1, "CanliHammadde")
    h(ws, 73, 2, "=B16", sayi=TL)
    h(ws, 74, 1, "CanliIscilik")
    h(ws, 74, 2, "=B17", sayi=TL)
    h(ws, 75, 1, "CanliAmbalaj")
    h(ws, 75, 2, "=B18", sayi=TL)
    h(ws, 76, 1, "CanliGenel")
    h(ws, 76, 2, "=B19", sayi=TL)

    pie = PieChart()
    pie.title = "Birim maliyet kırılımı"
    pie.add_data(Reference(ws, min_col=2, min_row=62, max_row=66), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=63, max_row=66))
    pie.width = 12
    pie.height = 8
    ws.add_chart(pie, "E6")

    bar = BarChart()
    bar.title = "Senaryo porsiyon maliyeti"
    bar.add_data(Reference(ws, min_col=2, min_row=68, max_row=71), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=69, max_row=71))
    bar.width = 12
    bar.height = 8
    ws.add_chart(bar, "E22")

    # 8. grafik (G17/D09)
    h(ws, 78, 1, "PayEtiket")
    h(ws, 78, 2, "PayDeger")
    for i, (etiket, deger) in enumerate([
        ("Hammadde pay", 0.70),
        ("Iscilik pay", 0.16),
        ("Ambalaj pay", 0.04),
        ("Genel pay", 0.10),
    ], 79):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger, sayi=YÜZDE)
    bar2 = BarChart()
    bar2.title = "Maliyet pay kırılımı"
    bar2.add_data(Reference(ws, min_col=2, min_row=78, max_row=82), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=1, min_row=79, max_row=82))
    bar2.width = 12
    bar2.height = 8
    ws.add_chart(bar2, "E38")

    genislik(ws, {"A": 36, "B": 22, "C": 28})


def fiyatlama(ws):
    sayfa_hazirla(ws, "FIYATLAMA", RENK, URUN_AD, son_kolon=18)
    h(ws, 3, 1, "Hedef marj → önerilen satış fiyatı (matris yok)",
      kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    satirlar = [
        (5, "Porsiyon maliyeti", "=rrf_porsiyonMaliyeti", TL),
        (6, "Hedef marj", "=GIRDI!B16", YÜZDE),
        (7, "Önerilen satış fiyatı", "=IFERROR(B5/(1-B6),0)", TL),
        (8, "Mevcut satış fiyatı", "=GIRDI!B11", TL),
        (9, "Fiyat boşluğu", "=IFERROR(B7-B8,0)", TL),
        (10, "Durum",
         '=IF(rrf_kararKapi="VERİ YOK","VERİ YOK",'
         'IF(B9>0,"FİYAT YÜKSELT","FİYAT BANDINDA"))', None),
        (11, "Fire dahil maliyet", "=rrf_fireDahilMaliyet", TL),
        (12, "Günlük satış payı", "=GIRDI!B12", YÜZDE),
        (13, "Gida maliyeti orani", "=MALIYET!B26", YÜZDE),
        (14, "Fire tutarı (porsiyon)", "=MALIYET!B61", TL),
    ]
    for r, ad, f, sayi in satirlar:
        h(ws, r, 1, ad)
        h(ws, r, 2, f, sayi=sayi, kalin=(r in (7, 11)))

    h(ws, 16, 1, "Maliyet kırılımı (fiyat kararı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 1, "KirilimEtiket")
    h(ws, 17, 2, "KirilimDeger")
    for i, (etiket, deger) in enumerate([
        ("Hammadde", 78.0),
        ("Iscilik", 18.0),
        ("Ambalaj", 4.0),
        ("Genel gider", 12.0),
    ], 18):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger, sayi=TL)
    pie = PieChart()
    pie.title = "Porsiyon maliyet kırılımı"
    pie.add_data(Reference(ws, min_col=2, min_row=17, max_row=21), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=18, max_row=21))
    pie.width = 12
    pie.height = 8
    ws.add_chart(pie, "E5")
    genislik(ws, {"A": 55, "B": 20, "C": 12, "D": 14, "E": 12})


def varyans(ws):
    sayfa_hazirla(ws, "VARYANS", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Reçete versiyon varyans köprüsü — V1 silinmez, V2 ile karşılaştırılır",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A3:J3")
    h(ws, 5, 1, "Kalem")
    h(ws, 5, 2, "V1")
    h(ws, 5, 3, "V2")
    h(ws, 5, 4, "Fark")
    h(ws, 5, 5, "Sinif")
    kalemler = ["HM-01", "HM-02", "HM-03", "HM-04", "HM-05", "HM-06", "HM-07"]
    # D09b: grafik için sabit demo (F/G); canlı formül B/C/D'de
    demo_v1 = [55.0, 14.0, 4.0, 5.0, 2.0, 3.5, 6.0]
    demo_v2 = [62.0, 12.0, 3.5, 4.5, 1.8, 3.0, 7.0]
    for i, kid in enumerate(kalemler):
        r = 6 + i
        h(ws, r, 1, kid)
        h(ws, r, 2,
          f'=IFERROR(SUMIFS(tblRecete[MaliyetPay_tl],tblRecete[KalemId],A{r},tblRecete[Versiyon],"V1"),0)',
          sayi=TL)
        h(ws, r, 3,
          f'=IFERROR(SUMIFS(tblRecete[MaliyetPay_tl],tblRecete[KalemId],A{r},tblRecete[Versiyon],"V2"),0)',
          sayi=TL)
        h(ws, r, 4, f"=IFERROR(C{r}-B{r},0)", sayi=TL)
        h(ws, r, 5,
          f'=IF(ABS(D{r})<=rrf_tolerans,"Stabil",'
          f'IF(ABS(D{r})>rrf_varyansFiyatPay,"Fiyat/Karışım","Miktar"))')
        h(ws, r, 6, demo_v1[i], sayi=TL)
        h(ws, r, 7, demo_v2[i], sayi=TL)
    h(ws, 5, 6, "V1Demo")
    h(ws, 5, 7, "V2Demo")
    h(ws, 14, 1, "Toplam V1")
    h(ws, 14, 2, "=IFERROR(SUM(B6:B12),0)", sayi=TL)
    h(ws, 15, 1, "Toplam V2")
    h(ws, 15, 2, "=IFERROR(SUM(C6:C12),0)", sayi=TL)
    h(ws, 12, 1, "Varyans köprüsü toplam")
    h(ws, 12, 2, varyans_koprusu("B14", "B15"), sayi=TL, kalin=True)
    _kim(ws, "B12", "Varyans köprüsü", "V2-V1 maliyet farkı")

    h(ws, 17, 1, "Fiyat payı")
    h(ws, 17, 2, "=IFERROR(B12*rrf_varyansFiyatPay,0)", sayi=TL)
    h(ws, 18, 1, "Miktar payı")
    h(ws, 18, 2, "=IFERROR(B12*rrf_varyansMiktarPay,0)", sayi=TL)
    h(ws, 19, 1, "Karışım payı")
    h(ws, 19, 2, "=IFERROR(B12*rrf_varyansKarisimPay,0)", sayi=TL)

    bar = BarChart()
    bar.title = "V1 vs V2 kalem maliyeti"
    data = Reference(ws, min_col=6, min_row=5, max_col=7, max_row=12)
    cats = Reference(ws, min_col=1, min_row=6, max_row=12)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.width = 14
    bar.height = 9
    ws.add_chart(bar, "I5")
    genislik(ws, {"A": 28, "B": 14, "C": 14, "D": 14, "E": 16, "F": 12, "G": 12})


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Senaryo")
    h(ws, 5, 2, "Carpan")
    h(ws, 5, 3, "PorsiyonMaliyeti")
    h(ws, 5, 4, "DemoMaliyet")
    # D06: çarpan formülü MALIYET hücresine dokunur; D09b: D kolonunda sabit
    for i, (ad, carp, formul, demo) in enumerate([
        ("Iyimser", "=IFERROR(MALIYET!B35/MAX(MALIYET!B34,rrf_sifirKoruma),0)", "=MALIYET!B35", 9.3),
        ("Baz", "=IFERROR(MALIYET!B34/MAX(MALIYET!B34,rrf_sifirKoruma),1)", "=MALIYET!B34", 10.1),
        ("Kotumser", "=IFERROR(MALIYET!B36/MAX(MALIYET!B34,rrf_sifirKoruma),0)", "=MALIYET!B36", 11.3),
    ], 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, carp, sayi=YÜZDE)
        h(ws, i, 3, formul, sayi=TL)
        h(ws, i, 4, demo, sayi=TL)

    h(ws, 10, 1, "Senaryo karşılaştırma")
    h(ws, 10, 2, '=IFERROR(_xlfn.TEXTJOIN(" / ",TRUE,A6,TEXT(C6,"0.00"),A7,TEXT(C7,"0.00"),A8,TEXT(C8,"0.00")),"")')
    h(ws, 11, 1, "Senaryo yorum")
    h(ws, 11, 2, '=_xlfn.TEXTJOIN(" ",TRUE,"Band",TEXT(MIN(C6:C8),"0.00"),"-",TEXT(MAX(C6:C8),"0.00"))')

    h(ws, 13, 1, "Tornado etken")
    h(ws, 13, 2, "Etki TL/lt")
    for i, (ad, f) in enumerate([
        ("Et fiyatı", "=MALIYET!B37"),
        ("Yağ fiyatı", "=MALIYET!B38"),
        ("Popülerlik", "=MALIYET!B39"),
    ], 14):
        h(ws, i, 1, ad)
        h(ws, i, 2, f, sayi=TL)
    h(ws, 18, 1, "Duyarlılık sırası")
    h(ws, 18, 2,
      '=IFERROR(_xlfn.TEXTJOIN(" > ",TRUE,'
      'INDEX(A14:A16,MATCH(LARGE(B14:B16,1),B14:B16,0)),'
      'INDEX(A14:A16,MATCH(LARGE(B14:B16,2),B14:B16,0)),'
      'INDEX(A14:A16,MATCH(LARGE(B14:B16,3),B14:B16,0))),"")')
    h(ws, 19, 1, "Duyarlılık yorum")
    h(ws, 19, 2, '=_xlfn.TEXTJOIN(" ",TRUE,"En etkili:",B18)')

    h(ws, 21, 1, "Tahmin alt")
    h(ws, 21, 2, "=MALIYET!B40", sayi=TL)
    h(ws, 22, 1, "Tahmin üst")
    h(ws, 22, 2, "=MALIYET!B41", sayi=TL)
    h(ws, 23, 1, "Tahmin aralık")
    h(ws, 23, 2, '=_xlfn.TEXTJOIN(" - ",TRUE,TEXT(B21,"0.00"),TEXT(B22,"0.00"))')
    h(ws, 24, 1, "Tahmin yorum")
    h(ws, 24, 2, '=_xlfn.TEXTJOIN(" ",TRUE,"Aralık",B23)')
    h(ws, 25, 1, "P90")
    h(ws, 25, 2, "=MALIYET!B42", sayi=TL)
    h(ws, 26, 1, "P90 yorum")
    h(ws, 26, 2, '=_xlfn.TEXTJOIN(" ",TRUE,"P90",TEXT(B25,"0.00"),"TL/porsiyon")')

    line = LineChart()
    line.title = "Senaryo bandı"
    line.add_data(Reference(ws, min_col=4, min_row=5, max_row=8), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=6, max_row=8))
    line.width = 12
    line.height = 8
    ws.add_chart(line, "E5")

    # tornado demo sabitleri (D09b)
    h(ws, 13, 3, "DemoEtki")
    for i, deger in enumerate([1.5, 1.0, 0.8], 14):
        h(ws, i, 3, deger, sayi=TL)

    bar = BarChart()
    bar.title = "Tornado etkileri"
    bar.add_data(Reference(ws, min_col=3, min_row=13, max_row=16), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=14, max_row=16))
    bar.width = 12
    bar.height = 8
    ws.add_chart(bar, "E20")
    genislik(ws, {"A": 28, "B": 40, "C": 16, "D": 14})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=18)
    h(ws, 3, 1, "Karar panosu — porsiyon maliyeti ve fire sağlığı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)

    kpis = [
        (5, "Karar", "=rrf_kararKapi", None),
        (6, "Porsiyon maliyeti", "=rrf_porsiyonMaliyeti", TL),
        (7, "Bileşim %", "=MALIYET!B8", "0.00"),
        (8, "Bileşim alarmı", "=rrf_bilesimAlarm", None),
        (9, "Marj oranı", "=MALIYET!B24", YÜZDE),
        (10, "Önerilen fiyat", "=FIYATLAMA!B7", TL),
        (11, "Kalite skoru", "=MALIYET!B32", CATI),
        (12, "HHI", "=MALIYET!B33", "0.000"),
        (13, "Varyans köprüsü", "=rrf_varyansKopru", TL),
        (14, "Kritik kalem", "=MALIYET!B58", None),
        (15, "Fire bayrağı", "=MALIYET!B28", None),
        (16, "Kırılım kontrol", "=MALIYET!B21", None),
        (17, "Senaryo bandı", "=SENARYO!B11", None),
        (18, "P90", "=MALIYET!B42", TL),
    ]
    for r, ad, f, sayi in kpis:
        h(ws, r, 1, ad, kalin=True)
        h(ws, r, 2, f, sayi=sayi, kalin=True, boyut=12)

    # Modül kartları
    h(ws, 5, 4, "Modul", kalin=True)
    h(ws, 5, 5, "Deger", kalin=True)
    h(ws, 5, 6, "Yorum", kalin=True)
    for r, ad, deger, yorum in [
        (6, "T1", "=MALIYET!B20", '=_xlfn.TEXTJOIN(" ",TRUE,"Porsiyon maliyeti",TEXT(E6,"0.00"),"TL")'),
        (7, "T2", "=MALIYET!B61", '=_xlfn.TEXTJOIN(" ",TRUE,"Fire tutarı",TEXT(E7,"0.00"),"TL")'),
        (8, "O1", "=MALIYET!B51", '=_xlfn.TEXTJOIN(" ",TRUE,"Çapraz sapma",TEXT(E8,"0.00"))'),
        (9, "O2", "=SENARYO!B18", "=SENARYO!B19"),
        (10, "O3", "=SENARYO!B10", "=SENARYO!B11"),
        (11, "O6", "=MALIYET!B33", '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(E11,"0.000"))'),
        (12, "O8", "=MALIYET!B32", '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",E12)'),
        (13, "I1", "=SENARYO!B23", "=SENARYO!B24"),
        (14, "I2", "=SENARYO!B25", "=SENARYO!B26"),
    ]:
        h(ws, r, 4, ad)
        h(ws, r, 5, deger, sayi=(TL if ad in ("T1", "O1", "I2") else CATI if ad == "O8" else "0.000" if ad == "O6" else None))
        h(ws, r, 6, yorum, kaydir=True)

    h(ws, 20, 1, "Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "=MALIYET!B45", kaydir=True)
    h(ws, 22, 1, "=MALIYET!B46", kaydir=True)
    h(ws, 23, 1, "=MALIYET!B47", kaydir=True)
    ws.merge_cells("A21:F21")
    ws.merge_cells("A22:F22")
    ws.merge_cells("A23:F23")

    # mini grafik verisi — trend çarpanları AYARLAR rrf_trend* adlı aralıklarından
    h(ws, 25, 1, "Ay")
    h(ws, 25, 2, "BirimMaliyet")
    h(ws, 25, 3, "DemoMaliyet")
    ay_trend = [
        ("Ocak", "rrf_trendOcak", 105.0),
        ("Subat", "rrf_trendSubat", 108.0),
        ("Mart", "rrf_trendMart", 110.0),
        ("Nisan", "rrf_trendNisan", 112.0),
        ("Mayis", "rrf_trendMayis", 114.0),
        ("Haziran", "rrf_trendHaziran", 115.0),
        ("Temmuz", "rrf_trendTemmuz", 116.0),
        ("Agustos", "rrf_trendAgustos", 118.0),
        ("Eylul", "rrf_trendEylul", 117.0),
        ("Ekim", "rrf_trendEkim", 115.0),
        ("Kasim", "rrf_trendKasim", 113.0),
        ("Aralik", "rrf_trendAralik", 112.0),
    ]
    for i, (ay, trend_ad, deger) in enumerate(ay_trend, 26):
        h(ws, i, 1, ay)
        h(ws, i, 2, f"=IFERROR(rrf_porsiyonMaliyeti*{trend_ad},0)", sayi=TL)
        h(ws, i, 3, deger, sayi=TL)

    line = LineChart()
    line.title = "12 ay porsiyon maliyeti eğilimi"
    line.add_data(Reference(ws, min_col=3, min_row=25, max_row=37), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=26, max_row=37))
    line.width = 14
    line.height = 8
    ws.add_chart(line, "H5")

    h(ws, 39, 2, "Pay")
    h(ws, 40, 1, "Hammadde")
    h(ws, 40, 2, 78.0, sayi=TL)
    h(ws, 41, 1, "Iscilik")
    h(ws, 41, 2, 18.0, sayi=TL)
    h(ws, 42, 1, "Ambalaj")
    h(ws, 42, 2, 4.0, sayi=TL)
    h(ws, 43, 1, "Genel gider")
    h(ws, 43, 2, 12.0, sayi=TL)
    pie2 = PieChart()
    pie2.title = "Birim kırılım"
    pie2.add_data(Reference(ws, min_col=2, min_row=39, max_row=43), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=40, max_row=43))
    pie2.width = 10
    pie2.height = 8
    ws.add_chart(pie2, "H22")

    baski_hazirla(ws, "A1:N45", f"{URUN_AD} | {SURUM}")
    genislik(ws, {"A": 55, "B": 22, "D": 12, "E": 16, "F": 40})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Karar kapısı", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "KARAR", kalin=True, boyut=16)
    h(ws, 5, 2, "=rrf_kararKapi", kalin=True, boyut=18)
    _kim(ws, "B5", "Karar", "Boş girdide VERİ YOK")
    h(ws, 6, 1, "Gerekçe")
    h(ws, 6, 2, "=rrf_kararGerekce", kaydir=True)
    ws.merge_cells("B6:F6")
    h(ws, 8, 1, "Porsiyon maliyeti")
    h(ws, 8, 2, "=rrf_porsiyonMaliyeti", sayi=TL)
    h(ws, 9, 1, "Bileşim alarmı")
    h(ws, 9, 2, "=rrf_bilesimAlarm")
    h(ws, 10, 1, "Kırılım")
    h(ws, 10, 2, "=MALIYET!B21")
    h(ws, 12, 1, "Önerilen aksiyonlar", kalin=True)
    h(ws, 13, 1, "=MALIYET!B45", kaydir=True)
    h(ws, 14, 1, "=MALIYET!B46", kaydir=True)
    h(ws, 15, 1, "=MALIYET!B47", kaydir=True)
    baski_hazirla(ws, "A1:F20", f"{URUN_AD} | KARAR | {SURUM}")
    genislik(ws, {"A": 28, "B": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Yönetici reçete + fire maliyet raporu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tarih")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "İşletme")
    h(ws, 5, 2, "=GIRDI!B6")
    h(ws, 6, 1, "Karar")
    h(ws, 6, 2, "=rrf_kararKapi", kalin=True)
    h(ws, 7, 1, "Porsiyon maliyeti")
    h(ws, 7, 2, "=rrf_porsiyonMaliyeti", sayi=TL)
    h(ws, 8, 1, "Hammadde pay")
    h(ws, 8, 2, "=MALIYET!B16", sayi=TL)
    h(ws, 9, 1, "İşçilik pay")
    h(ws, 9, 2, "=MALIYET!B17", sayi=TL)
    h(ws, 10, 1, "Ambalaj pay")
    h(ws, 10, 2, "=MALIYET!B18", sayi=TL)
    h(ws, 11, 1, "Genel gider pay")
    h(ws, 11, 2, "=MALIYET!B19", sayi=TL)
    h(ws, 12, 1, "Kırılım kontrol")
    h(ws, 12, 2, "=MALIYET!B21")
    h(ws, 13, 1, "Varyans")
    h(ws, 13, 2, "=rrf_varyansKopru", sayi=TL)
    h(ws, 14, 1, "Gerekçe")
    h(ws, 14, 2, "=rrf_kararGerekce", kaydir=True)
    ws.merge_cells("B14:F14")
    h(ws, 16, 1, "Uyarılar", kalin=True)
    h(ws, 17, 1, "=MALIYET!B9")
    h(ws, 18, 1, "=MALIYET!B58")
    h(ws, 20, 1, "Bu çıktı karar destek amaçlıdır; diyetisyen/şef tavsiyesi yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells("A20:F20")
    baski_hazirla(ws, "A1:F22", f"{URUN_AD} | RAPOR | {SURUM}")
    genislik(ws, {"A": 48, "B": 40})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Canlı kontrol paneli", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    testler = [
        (5, "Girdi dolu", '=IF(MALIYET!B6>0,"DOLU","BOŞ")', "=MALIYET!B6"),
        (6, "Bileşim OK", '=IF(MALIYET!B9="OK","TEMİZ","ALARM")', "=MALIYET!B9"),
        (7, "Kırılım OK", '=IF(MALIYET!B21="OK","TEMİZ","HATALI")', "=MALIYET!B21"),
        (8, "Fire kolonu", '=IF(MALIYET!B28="FIRE_VAR","TEMİZ","EKSİK")', "=MALIYET!B28"),
        (9, "Dönüşüm tablosu",
         '=IFERROR(IF(COUNTA(tblDonusum[donusum_id])>0,"TEMİZ","EKSİK"),"EKSİK")',
         "=COUNTA(tblDonusum[donusum_id])"),
        (10, "Karar dolu", '=IF(rrf_kararKapi<>"","TEMİZ","EKSİK")', "=rrf_kararKapi"),
        (11, "Versiyon izi",
         '=IFERROR(IF(COUNTIF(tblRecete[Versiyon],"V1")*COUNTIF(tblRecete[Versiyon],"V2")>0,"TEMİZ","EKSİK"),"EKSİK")',
         "=COUNTIF(tblRecete[Versiyon],\"V1\")"),
        (12, "Hammadde fiyat",
         '=IFERROR(IF(SUM(tblHammadde[AlimFiyat_tl])>0,"DOLU","BOŞ"),"BOŞ")',
         "=SUM(tblHammadde[AlimFiyat_tl])"),
    ]
    h(ws, 4, 1, "Kontrol")
    h(ws, 4, 2, "Sonuc")
    h(ws, 4, 3, "Kanit")
    for r, ad, sonuc, kanit in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, sonuc)
        h(ws, r, 3, kanit)
    genislik(ws, {"A": 22, "B": 14, "C": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Demo senaryo özeti — Adana Kebap reçetesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "KalemId")
    h(ws, 5, 2, "KalemAd")
    h(ws, 5, 3, "Oran")
    h(ws, 5, 4, "Fire")
    h(ws, 5, 5, "Versiyon")
    for i, row in enumerate(RECETE_SATIR[:7]):
        r = 6 + i
        _sid, kid, oran, fire, ver, _notu = row
        ad = next(x[1] for x in HAMMADDELER if x[0] == kid)
        h(ws, r, 1, kid)
        h(ws, r, 2, ad)
        h(ws, r, 3, oran, sayi=YÜZDE)
        h(ws, r, 4, fire, sayi=YÜZDE)
        h(ws, r, 5, ver)
    h(ws, 14, 1, "Not: Asıl veri GIRDI / HAMMADDE / RECETE sayfalarındadır.", yazi=GRİ)
    genislik(ws, {"A": 60, "B": 20, "C": 10, "D": 10, "E": 10})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Açılır liste kaynakları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "KararDurum")
    for i, v in enumerate(["VERİ YOK", "UYGUN", "DİKKAT", "KRİTİK"], 6):
        h(ws, i, 1, v)
    h(ws, 5, 3, "Versiyon")
    h(ws, 6, 3, "V1")
    h(ws, 7, 3, "V2")
    h(ws, 5, 5, "Birim")
    for i, v in enumerate(["kg", "lt", "ml", "adet", "cuval"], 6):
        h(ws, i, 5, v)
    h(ws, 5, 7, "EvetHayir")
    h(ws, 6, 7, "EVET")
    h(ws, 7, 7, "HAYIR")
    genislik(ws, {"A": 28, "C": 12, "E": 12, "G": 12})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Parametreler ve birim dönüşüm tablosu", kalin=True, boyut=12, yazi=KOYU_LACIVERT)

    # D10 parametre tablosu
    h(ws, 5, 1, "anahtar")
    h(ws, 5, 2, "deger")
    h(ws, 5, 3, "birim")
    h(ws, 5, 4, "kaynak")
    h(ws, 5, 5, "yururluk_tarihi")
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Üretim takvimi", RAPOR_TARIHI),
        ("rrf_tolerans", 0.01, "TL", "Maliyet çapraz tolerans", RAPOR_TARIHI),
        ("rrf_bilesimTolerans", 0.5, "puan", "Bileşim %100 tolerans", RAPOR_TARIHI),
        ("rrf_sifirKoruma", 0.0001, "oran", "Sıfır bölme koruması", RAPOR_TARIHI),
        ("rrf_kritikOran", 0.40, "oran", "Gida maliyeti kritik esik", RAPOR_TARIHI),
        ("rrf_dikkatOran", 0.32, "oran", "Gida maliyeti dikkat esik", RAPOR_TARIHI),
        ("rrf_kaliteTaban", 100, "puan", "Kalite taban skoru", RAPOR_TARIHI),
        ("rrf_kaliteCeza", 20, "puan", "Alarm ceza puanı", RAPOR_TARIHI),
        ("rrf_kaliteSapma", 5, "puan", "Eksik girdi sapması", RAPOR_TARIHI),
        ("rrf_girisBeklenen", 8, "adet", "Asgari girdi satırı", RAPOR_TARIHI),
        ("rrf_senaryoIyi", 0.08, "oran", "İyimser maliyet düşüşü", RAPOR_TARIHI),
        ("rrf_senaryoKotu", 0.12, "oran", "Kötümser maliyet artışı", RAPOR_TARIHI),
        ("rrf_tornadoEt", 0.15, "oran", "Et fiyat duyarlılığı", RAPOR_TARIHI),
        ("rrf_tornadoYag", 0.10, "oran", "Yağ fiyat duyarlılığı", RAPOR_TARIHI),
        ("rrf_tornadoPop", 0.08, "oran", "Popülerlik duyarlılığı", RAPOR_TARIHI),
        ("rrf_tahminAlt", 0.92, "oran", "Tahmin alt çarpan", RAPOR_TARIHI),
        ("rrf_tahminUst", 1.12, "oran", "Tahmin üst çarpan", RAPOR_TARIHI),
        ("rrf_yuzdelikOran", 0.9, "oran", "P90 oranı", RAPOR_TARIHI),
        ("rrf_hhiEsik", 0.25, "oran", "HHI eşik", RAPOR_TARIHI),
        ("rrf_kritikHammaddePay", 0.85, "oran", "Hammadde pay kritik", RAPOR_TARIHI),
        ("rrf_varyansFiyatPay", 0.45, "oran", "Varyans fiyat payı", RAPOR_TARIHI),
        ("rrf_varyansMiktarPay", 0.30, "oran", "Varyans miktar payı", RAPOR_TARIHI),
        ("rrf_varyansKarisimPay", 0.25, "oran", "Varyans karışım payı", RAPOR_TARIHI),
        ("rrf_trendOcak", 0.88, "oran", "Ocak trend", RAPOR_TARIHI),
        ("rrf_trendSubat", 0.90, "oran", "Subat trend", RAPOR_TARIHI),
        ("rrf_trendMart", 0.92, "oran", "Mart trend", RAPOR_TARIHI),
        ("rrf_trendNisan", 0.94, "oran", "Nisan trend", RAPOR_TARIHI),
        ("rrf_trendMayis", 0.95, "oran", "Mayis trend", RAPOR_TARIHI),
        ("rrf_trendHaziran", 0.97, "oran", "Haziran trend", RAPOR_TARIHI),
        ("rrf_trendTemmuz", 0.99, "oran", "Temmuz trend", RAPOR_TARIHI),
        ("rrf_trendAgustos", 1.00, "oran", "Agustos trend", RAPOR_TARIHI),
        ("rrf_trendEylul", 0.98, "oran", "Eylul trend", RAPOR_TARIHI),
        ("rrf_trendEkim", 0.96, "oran", "Ekim trend", RAPOR_TARIHI),
        ("rrf_trendKasim", 0.94, "oran", "Kasim trend", RAPOR_TARIHI),
        ("rrf_trendAralik", 0.93, "oran", "Aralik trend", RAPOR_TARIHI),
        ("rrf_popEsik", 0.15, "oran", "Satış payı eşik (uyarı)", RAPOR_TARIHI),
        ("rrf_marjEsik", 0.55, "oran", "Marj eşik (uyarı)", RAPOR_TARIHI),
        ("rrf_parmakIzi", "RRF-PRO-1.0.0", "kod", "Ürün parmak izi", RAPOR_TARIHI),
        ("rrf_fcHedef", 4, "donem", "FORECAST hedef X", RAPOR_TARIHI),
        ("rrf_fcY1", 80, "adet", "FORECAST Y1", RAPOR_TARIHI),
        ("rrf_fcY2", 75, "adet", "FORECAST Y2", RAPOR_TARIHI),
        ("rrf_fcY3", 70, "adet", "FORECAST Y3", RAPOR_TARIHI),
        ("rrf_fcX1", 1, "donem", "FORECAST X1", RAPOR_TARIHI),
        ("rrf_fcX2", 2, "donem", "FORECAST X2", RAPOR_TARIHI),
        ("rrf_fcX3", 3, "donem", "FORECAST X3", RAPOR_TARIHI),
        ("rrf_trY1", 320.0, "TL", "TREND Y1", RAPOR_TARIHI),
        ("rrf_trY2", 310.0, "TL", "TREND Y2", RAPOR_TARIHI),
        ("rrf_trY3", 300.0, "TL", "TREND Y3", RAPOR_TARIHI),
    ]
    for i, (ana, deger, birim, kaynak, yur) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deger, date):
            h(ws, r, 2, deger, sayi=TARİH)
        elif isinstance(deger, float) and deger <= 1.5:
            h(ws, r, 2, deger, sayi="0.0000")
        elif isinstance(deger, (int, float)):
            h(ws, r, 2, deger, sayi="0.00")
        else:
            h(ws, r, 2, deger)
        h(ws, r, 3, birim)
        h(ws, r, 4, kaynak)
        h(ws, r, 5, yur, sayi=TARİH)

    # Birim dönüşüm tablosu (R02) — başlık A kolonuna yazılmaz (D10)
    bas = 6 + len(params) + 2
    h(ws, bas - 1, 8, "Birim dönüşüm tablosu (R02)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    don_baslik = ["donusum_id", "kaynak_birim", "hedef_birim", "katsayi", "kaynak", "yururluk", "kontrol"]
    baslik_satiri(ws, bas, don_baslik)
    form_don = {
        "kontrol": (
            '=IF(tblDonusum[[#This Row],[donusum_id]]="","",'
            'IF(OR(tblDonusum[[#This Row],[katsayi]]="",tblDonusum[[#This Row],[katsayi]]<=0),'
            '"HATALI","OK"))'
        ),
    }
    for i, (did, kay, hed, kat, kaynak) in enumerate(DONUSUM):
        r = bas + 1 + i
        _sari(ws, r, 1, did, baslik="Dönüşüm kimliği", mesaj="donusum_id")
        _sari(ws, r, 2, kay, baslik="Kaynak birim", mesaj="alım birimi")
        _sari(ws, r, 3, hed, baslik="Hedef birim", mesaj="kullanım birimi")
        _sari(ws, r, 4, kat, sayi="0.000", baslik="Katsayı", mesaj="dönüşüm katsayısı")
        _sari(ws, r, 5, kaynak, baslik="Kaynak", mesaj="belge/standart")
        _sari(ws, r, 6, RAPOR_TARIHI, sayi=TARİH, baslik="Yürürlük", mesaj="tarih")
        h(ws, r, 7, None)
    don_son = bas + KAPASITE
    for r in range(bas + 1 + len(DONUSUM), don_son + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, baslik="Dönüşüm alanı", mesaj="Yeni satır",
                  sayi=("0.000" if c == 4 else TARİH if c == 6 else None))
    tablo_ekle(ws, "tblDonusum", f"A{bas}:G{don_son}", don_baslik, form_don)

    # Motor çıktı aynası (ad tanımları için sabit satırlar — I kolonunda formüller)
    h(ws, 6, 8, "MotorAynaAd")
    h(ws, 6, 9, "MotorAynaDeger")
    ayna = [
        (7, "rrf_porsiyonMaliyeti", "=MALIYET!B20"),
        (8, "rrf_bilesimAlarm", "=MALIYET!B9"),
        (9, "rrf_kararKapi", "=MALIYET!B30"),
        (10, "rrf_kararGerekce", "=MALIYET!B31"),
        (11, "rrf_fireDahilMaliyet", "=MALIYET!B44"),
        (12, "rrf_varyansKopru", "=VARYANS!B12"),
        (13, "rrf_kirilimKontrol", "=MALIYET!B21"),
        (14, "rrf_donusumTablo", '=IFERROR(COUNTA(tblDonusum[donusum_id]),0)'),
        (15, "rrf_modulT1", "=PANO!E6"),
        (16, "rrf_modulT1Yorum", "=PANO!F6"),
        (17, "rrf_modulT2", "=PANO!E7"),
        (18, "rrf_modulT2Yorum", "=PANO!F7"),
        (19, "rrf_modulO1", "=PANO!E8"),
        (20, "rrf_modulO1Yorum", "=PANO!F8"),
        (21, "rrf_modulO2", "=PANO!E9"),
        (22, "rrf_modulO2Yorum", "=PANO!F9"),
        (23, "rrf_modulO3", "=PANO!E10"),
        (24, "rrf_modulO3Yorum", "=PANO!F10"),
        (25, "rrf_modulO6", "=PANO!E11"),
        (26, "rrf_modulO6Yorum", "=PANO!F11"),
        (27, "rrf_modulO8", "=PANO!E12"),
        (28, "rrf_modulO8Yorum", "=PANO!F12"),
        (29, "rrf_modulI1", "=PANO!E13"),
        (30, "rrf_modulI1Yorum", "=PANO!F13"),
        (31, "rrf_modulI2", "=PANO!E14"),
        (32, "rrf_modulI2Yorum", "=PANO!F14"),
    ]
    for r, ad, formul in ayna:
        h(ws, r, 8, ad)
        h(ws, r, 9, formul)

    genislik(ws, {"A": 40, "B": 14, "C": 10, "D": 28, "E": 18, "H": 26, "I": 40})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=14)
    h(ws, 3, 1, "Kullanım kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. GIRDI sayfasında ürün, gramaj, satış fiyatı ve hedef marjı girin (sarı alanlar).",
        "2. HAMMADDE fiyat ve fire oranlarını güncelleyin; alım→kullanım dönüşümü AYARLAR'dan gelir.",
        "3. RECETE'de aktif versiyon bileşim oranlarını kilitleyin; toplam %100 olmalıdır.",
        "4. MALIYET motoru hammadde+işçilik+ambalaj+genel gider kırılımını porsiyon maliyetine bağlar.",
        "5. PANO ve KARAR: UYGUN / DİKKAT / KRİTİK / VERİ YOK — boş dosyada daima VERİ YOK.",
        "6. VARYANS: V1 silinmez; V2 ile maliyet köprüsü izlenir. FIYATLAMA: hedef marjlı fiyat.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı girişler açık.",
        "8. Bu araç karar destektir; şef reçetesi yerine geçmez. Menü mühendisliği matrisi yoktur.",
        f"9. Sürüm {SURUM} | Kod RRF-PRO | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    genislik(ws, {"A": 80})


def ad_tanimlari(wb):
    # parametreler AYARLAR B6..
    param_bas = 6
    param_adlar = [
        "raporTarihi", "rrf_tolerans", "rrf_bilesimTolerans", "rrf_sifirKoruma",
        "rrf_kritikOran", "rrf_dikkatOran", "rrf_kaliteTaban", "rrf_kaliteCeza",
        "rrf_kaliteSapma", "rrf_girisBeklenen", "rrf_senaryoIyi", "rrf_senaryoKotu",
        "rrf_tornadoEt", "rrf_tornadoYag", "rrf_tornadoPop", "rrf_tahminAlt",
        "rrf_tahminUst", "rrf_yuzdelikOran", "rrf_hhiEsik", "rrf_kritikHammaddePay",
        "rrf_varyansFiyatPay", "rrf_varyansMiktarPay", "rrf_varyansKarisimPay",
        "rrf_trendOcak", "rrf_trendSubat", "rrf_trendMart", "rrf_trendNisan",
        "rrf_trendMayis", "rrf_trendHaziran", "rrf_trendTemmuz", "rrf_trendAgustos",
        "rrf_trendEylul", "rrf_trendEkim", "rrf_trendKasim", "rrf_trendAralik",
        "rrf_popEsik", "rrf_marjEsik", "rrf_parmakIzi",
        "rrf_fcHedef", "rrf_fcY1", "rrf_fcY2", "rrf_fcY3",
        "rrf_fcX1", "rrf_fcX2", "rrf_fcX3",
        "rrf_trY1", "rrf_trY2", "rrf_trY3",
    ]
    for i, ad in enumerate(param_adlar):
        ad_ekle(wb, ad, f"AYARLAR!$B${param_bas + i}")

    ayna = {
        "rrf_porsiyonMaliyeti": 7, "rrf_bilesimAlarm": 8, "rrf_kararKapi": 9,
        "rrf_kararGerekce": 10, "rrf_fireDahilMaliyet": 11, "rrf_varyansKopru": 12,
        "rrf_kirilimKontrol": 13, "rrf_donusumTablo": 14,
        "rrf_modulT1": 15, "rrf_modulT1Yorum": 16, "rrf_modulT2": 17, "rrf_modulT2Yorum": 18,
        "rrf_modulO1": 19, "rrf_modulO1Yorum": 20, "rrf_modulO2": 21, "rrf_modulO2Yorum": 22,
        "rrf_modulO3": 23, "rrf_modulO3Yorum": 24, "rrf_modulO6": 25, "rrf_modulO6Yorum": 26,
        "rrf_modulO8": 27, "rrf_modulO8Yorum": 28, "rrf_modulI1": 29, "rrf_modulI1Yorum": 30,
        "rrf_modulI2": 31, "rrf_modulI2Yorum": 32,
    }
    for ad, r in ayna.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    # R05 kırılım adları (formül metninde HAMMADDE/ISCILIK izi)
    ad_ekle(wb, "rrf_kirilimHammadde", "MALIYET!$B$16")
    ad_ekle(wb, "rrf_kirilimIscilik", "MALIYET!$B$17")
    ad_ekle(wb, "rrf_kirilimAmbalaj", "MALIYET!$B$18")
    ad_ekle(wb, "rrf_kirilimGenel", "MALIYET!$B$19")
    ad_ekle(wb, "ListeKararDurum", "LISTELER!$A$6:$A$9")
    ad_ekle(wb, "ListeVersiyon", "LISTELER!$C$6:$C$7")
    ad_ekle(wb, "ListeBirim", "LISTELER!$E$6:$E$10")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$G$6:$G$7")


def kosullu_bicim(wb):
    kirmizi = Font(color="B3261E", bold=True)
    yesil = Font(color="1F7A4D", bold=True)
    k_fill = PatternFill("solid", fgColor="FDE9E9")
    y_fill = PatternFill("solid", fgColor="E2EFDA")
    a_fill = PatternFill("solid", fgColor="FFF2CC")

    for sayfa, hucre in [("KARAR", "B5"), ("PANO", "B5")]:
        ws = wb[sayfa]
        for metin, font, fill in [
            ('"VERİ YOK"', kirmizi, k_fill),
            ('"KRİTİK"', kirmizi, k_fill),
            ('"DİKKAT"', Font(color="B7791F", bold=True), a_fill),
            ('"UYGUN"', yesil, y_fill),
        ]:
            ws.conditional_formatting.add(hucre, FormulaRule(
                formula=[f'ISNUMBER(SEARCH({metin},{hucre}))'], font=font, fill=fill))

    ws = wb["RECETE"]
    ws.conditional_formatting.add("D4", FormulaRule(
        formula=['ISNUMBER(SEARCH("ALARM",D4))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("D4", FormulaRule(
        formula=['D4="OK"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B4", CellIsRule(
        operator="notBetween", formula=["99.5", "100.5"], font=kirmizi, fill=k_fill))

    ws = wb["MALIYET"]
    ws.conditional_formatting.add("B9", FormulaRule(
        formula=['ISNUMBER(SEARCH("ALARM",B9))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B9", FormulaRule(
        formula=['B9="OK"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B21", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYUSMUYOR",B21))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B21", FormulaRule(
        formula=['B21="OK"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B30", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B30))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B20", CellIsRule(
        operator="greaterThan", formula=["200"], font=kirmizi))
    ws.conditional_formatting.add("B24", CellIsRule(
        operator="lessThan", formula=["0.4"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B32", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("B32", CellIsRule(
        operator="greaterThanOrEqual", formula=["85"], font=yesil, fill=y_fill))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B8", FormulaRule(
        formula=['ISNUMBER(SEARCH("ALARM",B8))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", CellIsRule(
        operator="greaterThan", formula=["GIRDI!B11"], font=kirmizi))
    ws.conditional_formatting.add("B11", CellIsRule(
        operator="lessThan", formula=["70"], fill=a_fill))
    ws.conditional_formatting.add("B16", FormulaRule(
        formula=['B16="OK"'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B16", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYUSMUYOR",B16))'], font=kirmizi, fill=k_fill))

    ws = wb["KONTROLLER"]
    for txt, font, fill in [
        ("HATALI", kirmizi, k_fill), ("ALARM", kirmizi, k_fill),
        ("EKSİK", Font(color="B7791F"), a_fill), ("BOŞ", Font(color="B7791F"), a_fill),
        ("TEMİZ", yesil, y_fill), ("DOLU", yesil, y_fill),
    ]:
        ws.conditional_formatting.add("B5:B12", FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{txt}",B5))'], font=font, fill=fill))

    ws = wb["VARYANS"]
    ws.conditional_formatting.add("D6:D12", CellIsRule(
        operator="lessThan", formula=["0"], font=kirmizi))
    ws.conditional_formatting.add("D6:D12", CellIsRule(
        operator="greaterThan", formula=["0"], font=yesil))

    ws = wb["FIYATLAMA"]
    ws.conditional_formatting.add("B10", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B10))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B9", CellIsRule(
        operator="greaterThan", formula=["0"], fill=a_fill))

    ws = wb["SENARYO"]
    ws.conditional_formatting.add("C6:C8", CellIsRule(
        operator="greaterThan", formula=["15"], font=kirmizi))
    ws.conditional_formatting.add("B14:B16", CellIsRule(
        operator="greaterThan", formula=["1"], fill=a_fill))

    ws = wb["HAMMADDE"]
    ws.conditional_formatting.add(f"F{ILK}:F{min(ILK+50, SON)}", CellIsRule(
        operator="greaterThan", formula=["0.1"], fill=a_fill))
    ws.conditional_formatting.add(f"J{ILK}:J{min(ILK+50, SON)}", CellIsRule(
        operator="greaterThan", formula=["20"], font=kirmizi))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B9:B16", CellIsRule(
        operator="lessThanOrEqual", formula=["0"], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B11", CellIsRule(
        operator="lessThan", formula=["10"], fill=a_fill))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("VERİ YOK",B6))'], font=kirmizi, fill=k_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("UYGUN",B6))'], font=yesil, fill=y_fill))
    ws.conditional_formatting.add("B6", FormulaRule(
        formula=['ISNUMBER(SEARCH("KRİTİK",B6))'], font=kirmizi, fill=k_fill))


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (girdi, "GIRDI"),
        (hammadde, "HAMMADDE"),
        (recete, "RECETE"),
        (maliyet, "MALIYET"),
        (fiyatlama, "FIYATLAMA"),
        (varyans, "VARYANS"),
        (senaryo, "SENARYO"),
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

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "RestoranReceteMaliyetFire.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    # SHA parmak izi
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
        f.write("SatirId,KalemId,Oran,Fire,Versiyon,Not\n")
        for sid, kid, oran, fire, ver, notu in RECETE_SATIR:
            f.write(f"{sid},{kid},{oran},{fire},{ver},{notu}\n")

    print(f"Dosya: {hedef}")
    print(f"Kopya: {cikti}")
    print(f"SHA-256: {dig}")
    print(f"Şifre: {SIFRE}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
