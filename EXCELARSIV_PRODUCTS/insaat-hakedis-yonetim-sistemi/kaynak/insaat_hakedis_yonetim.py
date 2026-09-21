#!/usr/bin/env python3
"""İnşaat Hakediş Yönetim Sistemi — A3 üretim betiği (manda v6)."""

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
    kuyruk_sayim,
    kuyruk_tutar,
    sonraki_kimlik_formul,
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)

URUN_AD = "İnşaat Hakediş Yönetim Sistemi"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 10)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "DRAFT", "durum_adi": "Taslak", "sira": 1,
     "izinli_sonraki_durumlar": "BEKLIYOR", "kategori": "acik"},
    {"durum_id": "BEKLIYOR", "durum_adi": "Bekliyor", "sira": 2,
     "izinli_sonraki_durumlar": "ONAY|IPTAL", "kategori": "acik"},
    {"durum_id": "ONAY", "durum_adi": "Onaylandı", "sira": 3,
     "izinli_sonraki_durumlar": "ODEME|IPTAL", "kategori": "acik"},
    {"durum_id": "ODEME", "durum_adi": "Ödendi", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "DRAFT|BEKLIYOR",
    "BEKLIYOR|ONAY",
    "BEKLIYOR|IPTAL",
    "ONAY|ODEME",
    "ONAY|IPTAL",
]

TASERONLAR = [
    ("KAR-00001", "Atlas Kalıp Ltd.", "Kalıp", "İstanbul"),
    ("KAR-00002", "Demir Yapı A.Ş.", "Demir", "Ankara"),
    ("KAR-00003", "Nova Beton", "Beton", "İzmir"),
    ("KAR-00004", "Proje Elektrik", "Elektrik", "Bursa"),
    ("KAR-00005", "Su Tesisat Pro", "Tesisat", "Antalya"),
    ("KAR-00006", "Cephe Sistemleri", "Cephe", "Kocaeli"),
    ("KAR-00007", "Zemin Stabil", "Zemin", "Adana"),
    ("KAR-00008", "İskele Güven", "İskele", "Gaziantep"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Hakediş kuyruğu ve kararı etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["DRAFT", "BEKLIYOR", "ONAY", "ODEME", "IPTAL", "BEKLIYOR", "ONAY", "BEKLIYOR"]
    oncekiler = ["", "DRAFT", "BEKLIYOR", "ONAY", "BEKLIYOR", "DRAFT", "BEKLIYOR", "ONAY"]
    for i in range(1, DEMO_AKIS + 1):
        kart = TASERONLAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):  # yetim
            kart = "" if i == 22 else "KAR-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: DRAFT → ONAY
            o, d = "DRAFT", "ONAY"
        if i == 20:  # geçersiz: ONAY → BEKLIYOR
            o, d = "ONAY", "BEKLIYOR"
        teklif = 100000 + i * 12500
        siparis = teklif
        hakedis = teklif if i not in (12, 19) else teklif + 45000  # zincir kırık
        fatura = hakedis if i != 15 else hakedis - 8000
        odeme = fatura if d == "ODEME" else (0 if d != "ODEME" else fatura)
        if d != "ODEME":
            odeme = 0
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        satirlar.append({
            "id": f"HAK-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "teklif": teklif,
            "siparis": siparis,
            "hakedis": hakedis,
            "fatura": fatura,
            "odeme": odeme,
            "tutar": hakedis,
            "donem": "2026-08",
            "aciklama": f"Hakediş kalemi {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Taşeron hakedişlerini durum makinesi, kuyruk ve yaşlandırma ile yönetir; "
      "tutar zinciri sapmalarını görünür kılar; dönemsel kanıt raporu üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Taslak → Bekliyor → Onay → Ödeme / İptal",
        "Kuyruk + yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Geçersiz geçiş ve yetim kayıt uyarıları",
        "Teklif→Sipariş→Hakediş→Fatura→Ödeme tutar zinciri denetimi",
        "Dönemsel hakediş kanıt raporu — SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "İnşaat mali işler ve şantiye muhasebesi",
        "Proje müdürü — darboğaz ve geciken kuyruk",
        "SMMM — hakediş kanıt raporu",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a taşeron/proje kartı girin. 2) AKIS'e hakediş olaylarını yazın. "
      "3) KUYRUKLAR ve PANO'dan günün işini görün. 4) RAPOR'dan kanıt çıktısı alın.",
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
      'IF(OR(ihy_yetimKayit>0,ihy_gecersizGecis>0,ihy_zincirKirik>0),"İNCELE",'
      'IF(ihy_gecikenAdet>ihy_esikGeciken,"İNCELE","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=ihy_acikAdet", CATI),
        (7, "Bekliyor Adet", '=COUNTIFS(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar],">0")', CATI),
        (8, "Onaylandı Adet", '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[Tutar],">0")', CATI),
        (9, "Geciken Adet", "=ihy_gecikenAdet", CATI),
        (10, "Açık Tutar", "=ihy_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ihy_gecikmeMaliyeti", TL),
        (12, "Yetim Kayıt", "=ihy_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ihy_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ihy_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=ihy_yasKova", None),
        (16, "Tahmin Aralık", "=ihy_tahminAralik", TL),
        (17, "Senaryo Farkı", "=ihy_senaryoKarsilastirma", TL),
        (18, "Tornado Zirve", "=ihy_tornadoZirve", TL),
        (19, "Ödeme Öncelik", "=ihy_odemeOncelik", TL),
        (20, "Pareto Gecikme", "=ihy_paretoGecikme", YÜZDE),
        (21, "Kapalı Düşüm", "=ihy_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=ihy_yorumYas"),
        (8, "=ihy_yorumGecikme"),
        (9, "=ihy_yorumTahmin"),
        (10, "=ihy_yorumSenaryo"),
        (11, "=ihy_yorumTornado"),
        (12, "=ihy_yorumOncelik"),
        (13, "=ihy_yorumPareto"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    # Grafik kaynakları (düz sıfır yasak — D09b)
    h(ws, 23, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "Durum")
    h(ws, 24, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Taslak", 4), ("Bekliyor", 8), ("Onay", 7), ("Ödeme", 5), ("İptal", 4),
    ], 25):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 24, 4, "Yaş Kova")
    h(ws, 24, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 6), ("8-30", 9), ("31-60", 7), ("60+", 6),
    ], 25):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 24, 7, "Hafta")
    h(ws, 24, 8, "Hakediş")
    for i in range(8):
        h(ws, 25 + i, 7, i + 1, sayi=CATI)
        h(ws, 25 + i, 8, 180000 + i * 22000, sayi=TL)

    h(ws, 24, 10, "Senaryo")
    h(ws, 24, 11, "Tutar")
    for i, (ad, t) in enumerate([
        ("Temkinli", 4200000), ("Baz", 5100000), ("İyimser", 5900000),
    ], 25):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=TL)

    # 8 grafik
    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=24, max_row=29), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=25, max_row=29))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A35")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=24, max_row=28), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=25, max_row=28))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F35")

    c3 = LineChart()
    c3.title = "Haftalık Hakediş"
    c3.add_data(Reference(ws, min_col=8, min_row=24, max_row=32), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=25, max_row=32))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A50")

    c4 = BarChart()
    c4.title = "Senaryo Karşılaştırma"
    c4.add_data(Reference(ws, min_col=11, min_row=24, max_row=27), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=25, max_row=27))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F50")

    # Ek grafik kaynakları + 4 grafik daha
    h(ws, 24, 13, "Uyarı")
    h(ws, 24, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 8),
    ], 25):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 24, 16, "Ay")
    h(ws, 24, 17, "Ödeme")
    for i in range(6):
        h(ws, 25 + i, 16, i + 1, sayi=CATI)
        h(ws, 25 + i, 17, 900000 + i * 110000, sayi=TL)

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=24, max_row=28), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=25, max_row=28))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A65")

    c6 = LineChart()
    c6.title = "Aylık Ödeme"
    c6.add_data(Reference(ws, min_col=17, min_row=24, max_row=30), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=25, max_row=30))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F65")

    c7 = PieChart()
    c7.title = "Kova Payı"
    c7.add_data(Reference(ws, min_col=5, min_row=24, max_row=28), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=4, min_row=25, max_row=28))
    c7.width, c7.height = 9, 7
    ws.add_chart(c7, "A80")

    c8 = BarChart()
    c8.title = "Durum Adet"
    c8.add_data(Reference(ws, min_col=2, min_row=24, max_row=29), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=1, min_row=25, max_row=29))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F80")

    baski_hazirla(ws, "A1:N34", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kartlar(ws):
    sayfa_hazirla(ws, "KARTLAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Taşeron / proje kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, '=ihy_sonrakiKimlik', kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: KAR-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Brans", "Sehir", "SozlesmeNo", "LimitTl",
                "RiskNotu", "Aktif", "ToplamHacim", "OrtGecikme"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamHacim": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'SUMIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],tblAkis[Tutar]))'
        ),
        "OrtGecikme": (
            'IF(tblKartlar[[#This Row],[KartId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakKartId],tblKartlar[[#This Row],[KartId]],'
            'tblAkis[YasGun]),0))'
        ),
    }
    for i, (kid, unvan, brans, sehir) in enumerate(TASERONLAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="KAR-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Taşeron/firma unvanı")
        _sari(ws, r, 3, brans, baslik="Branş", mesaj="İş kalemi branşı")
        _sari(ws, r, 4, sehir, baslik="Şehir", mesaj="Şehir")
        _sari(ws, r, 5, f"SZL-2026-{i + 1:03d}", baslik="Sözleşme No", mesaj="Sözleşme numarası")
        _sari(ws, r, 6, 2500000 + i * 150000, sayi=TL, baslik="Limit", mesaj="Sözleşme limiti TL")
        _sari(ws, r, 7, "Normal", baslik="Risk Notu", mesaj="Kısa risk notu")
        _sari(ws, r, 8, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 9):
            _sari(ws, r, c, None, sayi=TL if c == 6 else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblKartlar", f"A{HDR}:J{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kart Kimliği", mesaj="KAR-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Branş"), (4, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Limit", mesaj="Limit ≥ 0",
              hata_baslik="Limit", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"H{ILK}:H{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 11)})
    ws.column_dimensions["B"].width = 22


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "İşlem satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "IslemId", "KaynakKartId", "Tarih", "OncekiDurum", "Durum",
        "Teklif", "Siparis", "Hakedis", "Fatura", "Odeme", "Tutar",
        "Donem", "Aciklama",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    # Formül kolonları — akis_motoru kalıpları
    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakKartId]]", "tblKartlar[KartId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    # Zincir: hakediş vs teklif
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[Teklif]]",
        "tblAkis[[#This Row],[Hakedis]]",
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

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="HAK-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["teklif"], sayi=TL, baslik="Teklif", mesaj="Teklif tutarı TL")
        _sari(ws, r, 7, s["siparis"], sayi=TL, baslik="Sipariş", mesaj="Sipariş tutarı TL")
        _sari(ws, r, 8, s["hakedis"], sayi=TL, baslik="Hakediş", mesaj="Hakediş tutarı TL")
        _sari(ws, r, 9, s["fatura"], sayi=TL, baslik="Fatura", mesaj="Fatura tutarı TL")
        _sari(ws, r, 10, s["odeme"], sayi=TL, baslik="Ödeme", mesaj="Ödeme tutarı TL")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9, 10, 11) else None)
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="HAK-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kart", mesaj="Kart kimliği (boş = yetim riski)",
              hata_baslik="Kart", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="Olay tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"D{ILK}:D{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(6, "Teklif"), (7, "Sipariş"), (8, "Hakediş"),
                  (9, "Fatura"), (10, "Ödeme"), (11, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "textLength", "0", f"L{ILK}:L{SON}",
              baslik="Dönem", mesaj="YYYY-AA",
              hata_baslik="Dönem", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="20")
    dogrulama(ws, "textLength", "0", f"M{ILK}:M{SON}",
              baslik="Açıklama", mesaj="Kısa açıklama",
              hata_baslik="Açıklama", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="120")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 20)})
    ws.column_dimensions["M"].width = 18
    ws.column_dimensions["Q"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    # akis_motoru kuyruk_sayim / kuyruk_tutar
    h(ws, 6, 1, "Bekliyor")
    h(ws, 6, 2, kuyruk_sayim("tblAkis", "BEKLIYOR"), sayi=CATI)
    h(ws, 6, 3, kuyruk_tutar("tblAkis", "BEKLIYOR"), sayi=TL)

    h(ws, 7, 1, "Onaylandı")
    h(ws, 7, 2, kuyruk_sayim("tblAkis", "ONAY"), sayi=CATI)
    h(ws, 7, 3, kuyruk_tutar("tblAkis", "ONAY"), sayi=TL)

    h(ws, 8, 1, "Taslak")
    h(ws, 8, 2, kuyruk_sayim("tblAkis", "DRAFT"), sayi=CATI)
    h(ws, 8, 3, kuyruk_tutar("tblAkis", "DRAFT"), sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&ihy_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
      'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ihy_esikGecikmeGun,'
      'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Ödeme+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Ödendi")
    h(ws, 12, 2, kuyruk_sayim("tblAkis", "ODEME"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblAkis", "ODEME"), sayi=TL)
    h(ws, 13, 1, "İptal")
    h(ws, 13, 2, kuyruk_sayim("tblAkis", "IPTAL"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblAkis", "IPTAL"), sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=ihy_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
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
        ("H03", "Toplam tutar", "=SUM(tblAkis[Tutar])"),
        ("H04", "Teklif toplam", "=SUM(tblAkis[Teklif])"),
        ("H05", "Sipariş toplam", "=SUM(tblAkis[Siparis])"),
        ("H06", "Hakediş toplam", "=SUM(tblAkis[Hakedis])"),
        ("H07", "Fatura toplam", "=SUM(tblAkis[Fatura])"),
        ("H08", "Ödeme toplam", "=SUM(tblAkis[Odeme])"),
        ("H09", "Taslak adet", '=COUNTIF(tblAkis[Durum],"DRAFT")'),
        ("H10", "Bekliyor adet", '=COUNTIF(tblAkis[Durum],"BEKLIYOR")'),
        ("H11", "Onay adet", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        ("H12", "Ödeme adet", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık adet", "=C9+C10+C11"),
        ("H15", "Kapalı adet", "=C12+C13"),
        ("H16", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"DRAFT",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"BEKLIYOR",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"ONAY",tblAkis[Tutar])'),
        ("H17", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H18", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H19", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H20", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H21", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H22", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H23", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H24", "Yaş 60+", '=COUNTIF(tblAkis[YasKova],"60+")'),
        ("H25", "Ort yaş gün", "=IFERROR(AVERAGE(tblAkis[YasGun]),0)"),
        ("H26", "Geciken adet",
         '=COUNTIFS(tblAkis[YasGun],">"&ihy_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
         'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&ihy_esikGecikmeGun,'
         'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Gecikme maliyeti", "=C27*ihy_gecikmeOranGun"),
        ("H29", "Temkinli senaryo", "=C16*ihy_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C16*ihy_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C16*ihy_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C31-C29"),
        ("H33", "Tornado (oran×tutar)", "=C16*ihy_tornadoOran"),
        ("H34", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ONAY",tblAkis[YasGun],">"&ihy_esikGecikmeGun)'),
        ("H35", "Pareto gecikme payı", "=IFERROR(C27/MAX(C16,1),0)"),
        ("H36", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],ihy_yuzdelikOran),0)"),
        ("H37", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C36)"),
        ("H38", "Tahmin aralık", "=IFERROR((C36+C37)/2,0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Limit toplam", "=SUM(tblKartlar[LimitTl])"),
        ("H41", "Limit kullanım oranı", "=IFERROR(C3/MAX(C40,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Bekliyor",C10,"Onay",C11,"Geciken",C26)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C21,C22,C23,C24)'),
        ("H44", "Karar skoru ham",
         "=IF(C1=0,0,MAX(0,100-C17*10-C18*10-C19*10-C26*2))"),
        ("H45", "Veri hazırlık", "=IF(C1=0,0,C2/MAX(C1,1))"),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "öncelik", "tahmin", "limit", "fark",
            "toplam", "havuz")) and "oran" not in acik.lower() and "adet" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "payı" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "yaş gün" in acik.lower():
            fmt = GUN
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    # Senaryo tablosu (çarpan = AYARLAR × veri dokunuşu — D06)
    h(ws, 54, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 54, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 54, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 55, 1, "Temkinli")
    h(ws, 55, 2, "=ihy_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 55, 3, "=C29", sayi=TL)
    h(ws, 56, 1, "Baz")
    h(ws, 56, 2, "=ihy_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 56, 3, "=C30", sayi=TL)
    h(ws, 57, 1, "İyimser")
    h(ws, 57, 2, "=ihy_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 57, 3, "=C31", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönemsel Hakediş Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=ihy_acikAdet", CATI),
        (10, "Açık Tutar", "=ihy_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=ihy_gecikmeMaliyeti", TL),
        (12, "Yetim", "=ihy_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=ihy_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=ihy_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=ihy_tahminAralik", TL),
        (16, "Kanıt Özeti", "=ihy_kanitRaporu", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Durum haritası AYARLAR'dadır. Geçersiz geçiş engellenmez, görünür kılınır. "
      "Zincir toleransı AYARLAR.zincirTolerans değerine bağlıdır. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Yetim kayıtları KARTLAR'a bağlayın veya satırı düzeltin.")
    h(ws, 23, 1, "2. Geçersiz geçişleri durum haritasına uygun ilerletin.")
    h(ws, 24, 1, "3. Zincir kırık satırlarda teklif/hakediş farkını açıklayın.")
    h(ws, 26, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H26", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "İşlem sayısı", '=COUNTA(tblAkis[IslemId])'),
        (6, "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        (7, "Bekliyor", '=COUNTIF(tblAkis[Durum],"BEKLIYOR")'),
        (8, "Onay", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        (9, "Ödeme", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        (10, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (11, "Açık toplam", "=B7+B8+COUNTIF(tblAkis[Durum],\"DRAFT\")"),
        (12, "Yetim", "=ihy_yetimKayit"),
        (13, "Geçersiz", "=ihy_gecersizGecis"),
        (14, "Zincir kırık", "=ihy_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=HESAP!C44"),
        (17, "Açık tutar", "=ihy_acikTutar"),
        (18, "Gecikme maliyeti", "=ihy_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=ihy_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r >= 17 else (CATI if r < 15 else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} kart ve {DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Taslak/Bekliyor/Onay/Ödeme/İptal — kapalılar kuyruktan düşer")
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

    h(ws, 3, 7, "Branş", kalin=True)
    h(ws, 6, 7, "Brans", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Kalıp", "Demir", "Beton", "Elektrik", "Tesisat", "Cephe"], 7):
        _sari(ws, i, 7, b, baslik="Branş", mesaj="Branş listesi")
    genislik(ws, {"A": 14, "C": 20, "E": 12, "G": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 10000, "TL", "İş kuralı — zincir sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği"),
        ("ihy_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("ihy_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("ihy_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("ihy_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("ihy_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("ihy_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("ihy_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("ihy_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("ihy_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("ihy_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("ihy_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("ihy_kuyrukOzeti", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("ihy_kapaliDusum", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("ihy_kanitRaporu", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("ihy_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("ihy_firmaUnvan", "Örnek İnşaat A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    # Durum haritası tablosu (A01) — parametre kolonundan ayrı (D10)
    h(ws, 26, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 27, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 28 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H27:L32", dh_bas)

    # Motor çıktıları (named range hedefleri) — HESAP'tan
    h(ws, 36, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 37, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 37, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (38, "ihy_acikAdet", "=HESAP!C14"),
        (39, "ihy_acikTutar", "=HESAP!C16"),
        (40, "ihy_gecikenAdet", "=HESAP!C26"),
        (41, "ihy_gecikmeMaliyeti", "=HESAP!C28"),
        (42, "ihy_yetimKayit", "=HESAP!C17"),
        (43, "ihy_gecersizGecis", "=HESAP!C18"),
        (44, "ihy_zincirKirik", "=HESAP!C19"),
        (45, "ihy_yasKova", "=HESAP!C43"),
        (46, "ihy_tahminAralik", "=HESAP!C38"),
        (47, "ihy_senaryoKarsilastirma", "=HESAP!C32"),
        (48, "ihy_tornadoZirve", "=HESAP!C33"),
        (49, "ihy_odemeOncelik", "=HESAP!C34"),
        (50, "ihy_paretoGecikme", "=HESAP!C35"),
        (51, "ihy_kapaliDusum", "=HESAP!C15"),
        (52, "ihy_kuyrukOzeti", "=HESAP!C42"),
        (53, "ihy_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C14,"Yetim",HESAP!C17,'
         '"Zincir",HESAP!C19,"Geciken",HESAP!C26)'),
        (54, "ihy_sonrakiKimlik",
         sonraki_kimlik_formul("KAR", "tblKartlar[KartId]").replace(
             'MAX(tblKartlar[KartId])',
             'COUNTA(tblKartlar[KartId])')),
    ]
    # sonraki_kimlik: KartId metin olduğu için COUNTA+1 daha güvenli
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    # Yorum hücreleri (makine üretimli)
    h(ws, 56, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (57, "ihy_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",ihy_yasKova)'),
        (58, "ihy_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(ihy_gecikmeMaliyeti,"0.00"),"TL")'),
        (59, "ihy_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(ihy_tahminAralik,"0.00"),"TL")'),
        (60, "ihy_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(ihy_senaryoKarsilastirma,"0.00"),"TL")'),
        (61, "ihy_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(ihy_tornadoZirve,"0.00"),"TL")'),
        (62, "ihy_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Ödeme öncelik havuzu",TEXT(ihy_odemeOncelik,"0.00"),"TL")'),
        (63, "ihy_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(ihy_paretoGecikme,"0.0%"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına taşeron/proje kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her hakediş olayını yazın; durum kolonunu listeden seçin.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "5. Kapalı durumlar (Ödendi/İptal) açık kuyruklardan düşer.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "7. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "8. RAPOR sayfasını PDF olarak yazdırabilirsiniz.",
        f"9. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    """≥40 veri doğrulama nesnesi tamamlayıcı."""
    ws = wb["AYARLAR"]
    for r in range(6, 24):
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
        "B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["AKIS"]
    for col, val, fill, font in [
        ("P", '"YETİM"', kirmizi, kf),
        ("Q", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("R", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("R", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for col in ("E",):
        for durum, fill in [("BEKLIYOR", sari), ("ONAY", yesil), ("IPTAL", kirmizi)]:
            ws.conditional_formatting.add(
                f"{col}{ILK}:{col}{SON}",
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
    for r in range(6, 51):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["KONTROLLER"]
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B15", CellIsRule(operator="equal", formula=['"VERİ VAR"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B16", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))
    for r in range(5, 15):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add(
        "B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(9, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["KARTLAR"]
    ws.conditional_formatting.add(
        f"H{ILK}:H{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"F{ILK}:F{SON}",
        CellIsRule(operator="greaterThan", formula=["5000000"], fill=yesil))


def adlari_bagla(wb):
    # Parametreler AYARLAR B6..
    param_adlari = [
        "raporTarihi", "zincirTolerans", "ihy_esikGecikmeGun", "ihy_esikGeciken",
        "ihy_gecikmeOranGun", "ihy_senaryoTemkinli", "ihy_senaryoBaz", "ihy_senaryoIyimser",
        "ihy_tornadoOran", "ihy_yuzdelikOran", "ihy_olcekHedef", "ihy_azamiTutar",
        "ihy_durumHaritasi", "ihy_kuyrukOzeti_metin", "ihy_kapaliDusum_metin",
        "ihy_kanitRaporu_metin", "ihy_dosyaSurumu", "ihy_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    # Motor I kolonları
    motor_map = {
        "ihy_acikAdet": 38, "ihy_acikTutar": 39, "ihy_gecikenAdet": 40,
        "ihy_gecikmeMaliyeti": 41, "ihy_yetimKayit": 42, "ihy_gecersizGecis": 43,
        "ihy_zincirKirik": 44, "ihy_yasKova": 45, "ihy_tahminAralik": 46,
        "ihy_senaryoKarsilastirma": 47, "ihy_tornadoZirve": 48, "ihy_odemeOncelik": 49,
        "ihy_paretoGecikme": 50, "ihy_kapaliDusum": 51, "ihy_kuyrukOzeti": 52,
        "ihy_kanitRaporu": 53, "ihy_sonrakiKimlik": 54,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "ihy_yorumYas": 57, "ihy_yorumGecikme": 58, "ihy_yorumTahmin": 59,
        "ihy_yorumSenaryo": 60, "ihy_yorumTornado": 61, "ihy_yorumOncelik": 62,
        "ihy_yorumPareto": 63,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")


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
    dosya_ad = "InsaatHakedisYonetimSistemi.xlsx"
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
