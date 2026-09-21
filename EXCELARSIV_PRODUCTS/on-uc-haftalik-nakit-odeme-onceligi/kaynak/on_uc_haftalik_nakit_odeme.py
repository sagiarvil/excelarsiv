#!/usr/bin/env python3
"""13 Haftalık Nakit + Ödeme Önceliği — A3 üretim betiği (manda v6)."""

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
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)

URUN_AD = "13 Haftalık Nakit + Ödeme Önceliği"
SURUM = "1.0.0"
RENK = "0B3D2E"
KAPASITE = 250
DEMO_KART = 8
DEMO_AKIS = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 10)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "PLAN", "durum_adi": "Planlandı", "sira": 1,
     "izinli_sonraki_durumlar": "ONAY|IPTAL", "kategori": "acik"},
    {"durum_id": "ONAY", "durum_adi": "Onaylandı", "sira": 2,
     "izinli_sonraki_durumlar": "ODEME|ERTELE|IPTAL", "kategori": "acik"},
    {"durum_id": "ERTELE", "durum_adi": "Ertelendi", "sira": 3,
     "izinli_sonraki_durumlar": "ONAY|IPTAL", "kategori": "acik"},
    {"durum_id": "ODEME", "durum_adi": "Ödendi", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "PLAN|ONAY",
    "PLAN|IPTAL",
    "ONAY|ODEME",
    "ONAY|ERTELE",
    "ONAY|IPTAL",
    "ERTELE|ONAY",
    "ERTELE|IPTAL",
]

ALACAKLILAR = [
    ("ALN-00001", "Atlas Tedarik Ltd.", "Tedarikçi", "İstanbul"),
    ("ALN-00002", "Demir Finans A.Ş.", "Kredi", "Ankara"),
    ("ALN-00003", "Nova Enerji", "Enerji", "İzmir"),
    ("ALN-00004", "Proje Personel", "Bordro", "Bursa"),
    ("ALN-00005", "Su Vergi Büro", "Vergi", "Antalya"),
    ("ALN-00006", "Cephe Kira A.Ş.", "Kira", "Kocaeli"),
    ("ALN-00007", "Zemin Sigorta", "Sigorta", "Adana"),
    ("ALN-00008", "İskele Lojistik", "Lojistik", "Gaziantep"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Ödeme kuyruğu ve nakit kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    """≥25 satır; yetim, geçersiz geçiş ve zincir kırığı kasıtlı."""
    satirlar = []
    durumlar = ["PLAN", "ONAY", "ERTELE", "ODEME", "IPTAL", "ONAY", "PLAN", "ERTELE"]
    oncekiler = ["", "PLAN", "ONAY", "ONAY", "ONAY", "PLAN", "", "ONAY"]
    oncelikler = ["KRİTİK", "YÜKSEK", "NORMAL", "KRİTİK", "NORMAL", "YÜKSEK", "KRİTİK", "NORMAL"]
    for i in range(1, DEMO_AKIS + 1):
        kart = ALACAKLILAR[(i - 1) % DEMO_KART][0]
        if i in (22, 25):  # yetim
            kart = "" if i == 22 else "ALN-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:  # geçersiz: PLAN → ODEME
            o, d = "PLAN", "ODEME"
        if i == 20:  # geçersiz: ODEME → ONAY
            o, d = "ODEME", "ONAY"
        planlanan = 80000 + i * 9500
        onaylanan = planlanan if i not in (12, 19) else planlanan + 42000  # zincir kırık
        odeme = onaylanan if d == "ODEME" else 0
        tar = RAPOR_TARIHI - timedelta(days=(i * 3) % 75)
        satirlar.append({
            "id": f"ODM-{i:05d}",
            "kart": kart,
            "tarih": tar,
            "onceki": o,
            "durum": d,
            "planlanan": planlanan,
            "onaylanan": onaylanan,
            "odeme": odeme,
            "tutar": onaylanan,
            "hafta": ((i - 1) % 13) + 1,
            "oncelik": oncelikler[(i - 1) % len(oncelikler)],
            "donem": "2026-W32",
            "aciklama": f"Ödeme kalemi {i}",
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "13 haftalık nakit çıkışlarını durum makinesi, ödeme önceliği ve yaşlandırma ile yönetir; "
      "plan–onay–ödeme tutar zincirini görünür kılar; haftalık kanıt raporu üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Planlandı → Onaylandı → Ertelendi → Ödeme / İptal",
        "Ödeme kuyruğu + yaşlandırma kovaları (0-7 / 8-30 / 31-60 / 60+)",
        "Geçersiz geçiş ve yetim kayıt uyarıları",
        "Planlanan→Onaylanan→Ödeme tutar zinciri + öncelik skoru",
        "13 haftalık nakit kanıt raporu — CFO / ortak sunumuna uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "CFO ve mali işler — haftalık ödeme önceliği",
        "Muhasebe — kuyruk güncelleme",
        "Ortak — nakit açığı görünürlüğü",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) KARTLAR'a alacaklı/tedarikçi kartı girin. 2) AKIS'e ödeme olaylarını yazın. "
      "3) KUYRUKLAR ve PANO'dan haftanın önceliğini görün. 4) RAPOR'dan kanıt çıktısı alın.",
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
      'IF(OR(onh_yetimKayit>0,onh_gecersizGecis>0,onh_zincirKirik>0),"İNCELE",'
      'IF(onh_gecikenAdet>onh_esikGeciken,"İNCELE","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=onh_acikAdet", CATI),
        (7, "Planlandı Adet", '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[Tutar],">0")', CATI),
        (8, "Onaylandı Adet", '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[Tutar],">0")', CATI),
        (9, "Ertelendi Adet", '=COUNTIFS(tblAkis[Durum],"ERTELE",tblAkis[Tutar],">0")', CATI),
        (10, "Geciken Adet", "=onh_gecikenAdet", CATI),
        (11, "Açık Tutar", "=onh_acikTutar", TL),
        (12, "Gecikme Maliyeti", "=onh_gecikmeMaliyeti", TL),
        (13, "Yetim Kayıt", "=onh_yetimKayit", CATI),
        (14, "Geçersiz Geçiş", "=onh_gecersizGecis", CATI),
        (15, "Zincir Kırık", "=onh_zincirKirik", CATI),
        (16, "Yaş Kova Özeti", "=onh_yasKova", None),
        (17, "Tahmin Aralık", "=onh_tahminAralik", TL),
        (18, "Senaryo Farkı", "=onh_senaryoKarsilastirma", TL),
        (19, "Tornado Zirve", "=onh_tornadoZirve", TL),
        (20, "Öncelik Skoru", "=onh_oncelikSkoru", CATI),
        (21, "Pareto Gecikme", "=onh_paretoGecikme", YÜZDE),
        (22, "Kapalı Düşüm", "=onh_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=onh_yorumYas"),
        (8, "=onh_yorumGecikme"),
        (9, "=onh_yorumTahmin"),
        (10, "=onh_yorumSenaryo"),
        (11, "=onh_yorumTornado"),
        (12, "=onh_yorumOncelik"),
        (13, "=onh_yorumPareto"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    # Grafik kaynakları (düz sıfır yasak — D09b)
    h(ws, 23, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1, "Durum")
    h(ws, 24, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Plan", 5), ("Onay", 7), ("Ertele", 6), ("Ödeme", 5), ("İptal", 5),
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
    h(ws, 24, 8, "Nakit Çıkış")
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
    c3.title = "Haftalık Nakit Çıkış"
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
    alt_bant(ws, 3, "Alacaklı / tedarikçi kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, '=onh_sonrakiKimlik', kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: ALN-00009 | Risk: Elle çakışma")

    sutunlar = ["KartId", "Unvan", "Tur", "Sehir", "SozlesmeNo", "LimitTl",
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
    for i, (kid, unvan, brans, sehir) in enumerate(ALACAKLILAR):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Kart Kimliği", mesaj="ALN-##### formatında")
        _sari(ws, r, 2, unvan, baslik="Unvan", mesaj="Taşeron/firma unvanı")
        _sari(ws, r, 3, brans, baslik="Tür", mesaj="Alacaklı türü")
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
              baslik="Kart Kimliği", mesaj="ALN-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Unvan"), (3, "Tür"), (4, "Şehir"), (5, "Sözleşme"), (7, "Risk")]:
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
        "Planlanan", "Onaylanan", "Odeme", "Tutar", "HaftaNo", "OncelikSinif",
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
        "tblAkis[[#This Row],[Planlanan]]",
        "tblAkis[[#This Row],[Onaylanan]]",
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
        _sari(ws, r, 1, s["id"], baslik="İşlem Kimliği", mesaj="ODM-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Kart", mesaj="KARTLAR'daki KartId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Olay tarihi")
        _sari(ws, r, 4, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 5, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 6, s["planlanan"], sayi=TL, baslik="Planlanan", mesaj="Planlanan tutar TL")
        _sari(ws, r, 7, s["onaylanan"], sayi=TL, baslik="Onaylanan", mesaj="Onaylanan tutar TL")
        _sari(ws, r, 8, s["odeme"], sayi=TL, baslik="Ödeme", mesaj="Ödenen tutar TL")
        _sari(ws, r, 9, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Kuyruk tutarı TL")
        _sari(ws, r, 10, s["hafta"], sayi=CATI, baslik="Hafta", mesaj="1-13 hafta numarası")
        _sari(ws, r, 11, s["oncelik"], baslik="Öncelik", mesaj="KRİTİK/YÜKSEK/NORMAL")
        _sari(ws, r, 12, s["donem"], baslik="Dönem", mesaj="YYYY-W##")
        _sari(ws, r, 13, s["aciklama"], baslik="Açıklama", mesaj="Kısa açıklama")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 14):
            fmt = TARİH if c == 3 else (TL if c in (6, 7, 8, 9) else (CATI if c == 10 else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:S{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="İşlem Kimliği", mesaj="ODM-##### girin",
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
    for c, ad in [(6, "Planlanan"), (7, "Onaylanan"), (8, "Ödeme"), (9, "Tutar")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "whole", "1", f"J{ILK}:J{SON}",
              baslik="Hafta", mesaj="1-13 hafta numarası",
              hata_baslik="Hafta", hata_mesaj="1 ile 13 arası",
              isaret="between", f2="13")
    dogrulama(ws, "list", "ListeOncelik", f"K{ILK}:K{SON}",
              baslik="Öncelik", mesaj="KRİTİK/YÜKSEK/NORMAL",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "textLength", "0", f"L{ILK}:L{SON}",
              baslik="Dönem", mesaj="YYYY-W##",
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
    h(ws, 6, 1, "Planlandı")
    h(ws, 6, 2, '=COUNTIFS(tblAkis[Durum],"PLAN",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 6, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"PLAN")', sayi=TL)

    h(ws, 7, 1, "Onaylandı")
    h(ws, 7, 2, '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 7, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ONAY")', sayi=TL)

    h(ws, 8, 1, "Ertelendi")
    h(ws, 8, 2, '=COUNTIFS(tblAkis[Durum],"ERTELE",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 8, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ERTELE")', sayi=TL)

    h(ws, 9, 1, "Geciken (açık, yaş>eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[YasGun],">"&onh_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
      'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&onh_esikGecikmeGun,'
      'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")',
      sayi=TL)

    h(ws, 11, 1, "Kapalı (Ödeme+İptal) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Ödendi")
    h(ws, 12, 2, '=COUNTIFS(tblAkis[Durum],"ODEME",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 12, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ODEME")', sayi=TL)
    h(ws, 13, 1, "İptal")
    h(ws, 13, 2, '=COUNTIFS(tblAkis[Durum],"IPTAL",tblAkis[Tutar],">0")', sayi=CATI)
    h(ws, 13, 3, '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"IPTAL")', sayi=TL)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=onh_kuyrukOzeti", kaydir=True)
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
        ("H04", "Planlanan toplam", "=SUM(tblAkis[Planlanan])"),
        ("H05", "Onaylanan toplam", "=SUM(tblAkis[Onaylanan])"),
        ("H06", "Ödeme toplam", "=SUM(tblAkis[Odeme])"),
        ("H07", "Plan-onay farkı", "=ABS(C10-C9)"),
        ("H08", "Ödeme oranı", "=IFERROR(C11/MAX(C10,1),0)"),
        ("H09", "Plan adet", '=COUNTIF(tblAkis[Durum],"PLAN")'),
        ("H10", "Onay adet", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        ("H11", "Ertele adet", '=COUNTIF(tblAkis[Durum],"ERTELE")'),
        ("H12", "Ödeme adet", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        ("H13", "İptal adet", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        ("H14", "Açık adet", "=C14+C15+C16"),
        ("H15", "Kapalı adet", "=C17+C18"),
        ("H16", "Açık tutar",
         '=SUMIF(tblAkis[Durum],"PLAN",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"ONAY",tblAkis[Tutar])'
         '+SUMIF(tblAkis[Durum],"ERTELE",tblAkis[Tutar])'),
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
         '=COUNTIFS(tblAkis[YasGun],">"&onh_esikGecikmeGun,tblAkis[Durum],"<>ODEME",'
         'tblAkis[Durum],"<>IPTAL",tblAkis[Tutar],">0")'),
        ("H27", "Geciken tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[YasGun],">"&onh_esikGecikmeGun,'
         'tblAkis[Durum],"<>ODEME",tblAkis[Durum],"<>IPTAL")'),
        ("H28", "Gecikme maliyeti", "=C32*onh_gecikmeOranGun"),
        ("H29", "Temkinli senaryo", "=C21*onh_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C21*onh_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C21*onh_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×tutar)", "=C21*onh_tornadoOran"),
        ("H34", "Ödeme öncelik havuzu",
         '=SUMIFS(tblAkis[Tutar],tblAkis[Durum],"ONAY",tblAkis[YasGun],">"&onh_esikGecikmeGun)'),
        ("H35", "Pareto gecikme payı", "=IFERROR(C32/MAX(C21,1),0)"),
        ("H36", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[Tutar],onh_yuzdelikOran),0)"),
        ("H37", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[IslemId])+1,tblAkis[Tutar],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "Kart sayısı", '=COUNTA(tblKartlar[KartId])'),
        ("H40", "Limit toplam", "=SUM(tblKartlar[LimitTl])"),
        ("H41", "Limit kullanım oranı", "=IFERROR(C8/MAX(C45,1),0)"),
        ("H42", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Plan",C14,"Onay",C15,"Ertele",C16,"Geciken",C31)'),
        ("H43", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C26,C27,C28,C29)'),
        ("H44", "Öncelik skoru",
         "=IF(C6=0,0,ROUND(C31*3+C22*5+C23*5+C24*5+C16*2+C15,0))"),
        ("H45", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C22*10-C23*10-C24*10-C31*2))"),
        ("H46", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H47", "Hafta 1-4 tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[HaftaNo],">=1",tblAkis[HaftaNo],"<=4")'),
        ("H48", "Hafta 5-8 tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[HaftaNo],">=5",tblAkis[HaftaNo],"<=8")'),
        ("H49", "Hafta 9-13 tutar",
         '=SUMIFS(tblAkis[Tutar],tblAkis[HaftaNo],">=9",tblAkis[HaftaNo],"<=13")'),
        ("H50", "Kritik öncelik tutar",
         '=SUMIF(tblAkis[OncelikSinif],"KRİTİK",tblAkis[Tutar])'),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "senaryo", "tornado", "tahmin", "limit", "fark",
            "toplam", "havuz")) and "oran" not in acik.lower() and "adet" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "payı" in acik.lower() or "hazırlık" in acik.lower()
            else (None if "metin" in acik.lower() or "özet" in acik.lower() else CATI)
        )
        if "yaş gün" in acik.lower():
            fmt = GUN
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    # Senaryo tablosu (çarpan = AYARLAR × veri dokunuşu — D06)
    h(ws, 58, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 58, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 58, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 59, 1, "Temkinli")
    h(ws, 59, 2, "=onh_senaryoTemkinli*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 59, 3, "=C34", sayi=TL)
    h(ws, 60, 1, "Baz")
    h(ws, 60, 2, "=onh_senaryoBaz*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 60, 3, "=C35", sayi=TL)
    h(ws, 61, 1, "İyimser")
    h(ws, 61, 2, "=onh_senaryoIyimser*IF(COUNTA(tblAkis[IslemId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 61, 3, "=C36", sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 50})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "13 Haftalık Nakit Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Açık Kuyruk", "=onh_acikAdet", CATI),
        (10, "Açık Tutar", "=onh_acikTutar", TL),
        (11, "Gecikme Maliyeti", "=onh_gecikmeMaliyeti", TL),
        (12, "Yetim", "=onh_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=onh_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=onh_zincirKirik", CATI),
        (15, "Tahmin Aralık", "=onh_tahminAralik", TL),
        (16, "Kanıt Özeti", "=onh_kanitRaporu", None),
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
    h(ws, 24, 1, "3. Zincir kırık satırlarda planlanan/onaylanan farkını açıklayın.")
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
        (8, "Onay", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        (9, "Ertele", '=COUNTIF(tblAkis[Durum],"ERTELE")'),
        (10, "Ödeme", '=COUNTIF(tblAkis[Durum],"ODEME")'),
        (11, "İptal", '=COUNTIF(tblAkis[Durum],"IPTAL")'),
        (12, "Açık toplam", "=B7+B8+B9"),
        (13, "Yetim", "=onh_yetimKayit"),
        (14, "Geçersiz", "=onh_gecersizGecis"),
        (15, "Zincir kırık", "=onh_zincirKirik"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=HESAP!C50"),
        (18, "Açık tutar", "=onh_acikTutar"),
        (19, "Gecikme maliyeti", "=onh_gecikmeMaliyeti"),
        (20, "Kapalı düşüm", "=onh_kapaliDusum"),
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
      f"Dosyada {DEMO_KART} kart ve {DEMO_AKIS} akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "KARTLAR ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Plan/Onay/Ertele/Ödeme/İptal — kapalılar kuyruktan düşer")
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

    h(ws, 3, 7, "Tür", kalin=True)
    h(ws, 6, 7, "Tur", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["Kalıp", "Demir", "Beton", "Elektrik", "Tesisat", "Cephe"], 7):
        _sari(ws, i, 7, b, baslik="Tür", mesaj="Branş listesi")
    h(ws, 3, 9, "Öncelik", kalin=True)
    h(ws, 6, 9, "Oncelik", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(["KRİTİK", "YÜKSEK", "NORMAL"], 7):
        _sari(ws, i, 9, b, baslik="Öncelik", mesaj="Öncelik sınıfı")
    genislik(ws, {"A": 14, "C": 20, "E": 12, "G": 14, "I": 12})


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
        ("onh_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("onh_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("onh_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("onh_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("onh_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("onh_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("onh_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("onh_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("onh_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("onh_azamiTutar", 50000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("onh_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("onh_kuyrukOzeti", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("onh_kapaliDusum", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("onh_kanitRaporu", "RAPOR kanıt çıktısı", "metin", "Dönemsel rapor", "01.01.2026",
         "Kanıt raporu"),
        ("onh_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("onh_firmaUnvan", "Örnek Holding A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma Unvanı"),
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
        (38, "onh_acikAdet", "=HESAP!C19"),
        (39, "onh_acikTutar", "=HESAP!C21"),
        (40, "onh_gecikenAdet", "=HESAP!C31"),
        (41, "onh_gecikmeMaliyeti", "=HESAP!C33"),
        (42, "onh_yetimKayit", "=HESAP!C22"),
        (43, "onh_gecersizGecis", "=HESAP!C23"),
        (44, "onh_zincirKirik", "=HESAP!C24"),
        (45, "onh_yasKova", "=HESAP!C48"),
        (46, "onh_tahminAralik", "=HESAP!C43"),
        (47, "onh_senaryoKarsilastirma", "=HESAP!C37"),
        (48, "onh_tornadoZirve", "=HESAP!C38"),
        (49, "onh_oncelikSkoru", "=HESAP!C49"),
        (50, "onh_paretoGecikme", "=HESAP!C40"),
        (51, "onh_kapaliDusum", "=HESAP!C20"),
        (52, "onh_kuyrukOzeti", "=HESAP!C47"),
        (53, "onh_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C19,"Yetim",HESAP!C22,'
         '"Zincir",HESAP!C24,"Geciken",HESAP!C31)'),
        (54, "onh_sonrakiKimlik",
         '=CONCATENATE("ALN-",TEXT(COUNTA(tblKartlar[KartId])+1,"00000"))'),
    ]
    # sonraki_kimlik: KartId metin olduğu için COUNTA+1 daha güvenli
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    # Yorum hücreleri (makine üretimli)
    h(ws, 56, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (57, "onh_yorumYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",onh_yasKova)'),
        (58, "onh_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(onh_gecikmeMaliyeti,"0.00"),"TL")'),
        (59, "onh_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(onh_tahminAralik,"0.00"),"TL")'),
        (60, "onh_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(onh_senaryoKarsilastirma,"0.00"),"TL")'),
        (61, "onh_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(onh_tornadoZirve,"0.00"),"TL")'),
        (62, "onh_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Öncelik skoru",TEXT(onh_oncelikSkoru,"0"),"puan")'),
        (63, "onh_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto gecikme payı",TEXT(onh_paretoGecikme,"0.0%"))'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. KARTLAR sayfasına alacaklı/tedarikçi kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her ödeme olayını yazın; durum kolonunu listeden seçin.",
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
    for r in range(7, 14):
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
    for r in range(6, 23):
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
        for durum, fill in [("PLAN", sari), ("ERTELE", sari), ("ONAY", yesil), ("IPTAL", kirmizi)]:
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
    for r in range(6, 56):
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
        "raporTarihi", "zincirTolerans", "onh_esikGecikmeGun", "onh_esikGeciken",
        "onh_gecikmeOranGun", "onh_senaryoTemkinli", "onh_senaryoBaz", "onh_senaryoIyimser",
        "onh_tornadoOran", "onh_yuzdelikOran", "onh_olcekHedef", "onh_azamiTutar",
        "onh_durumHaritasi", "onh_kuyrukOzeti_metin", "onh_kapaliDusum_metin",
        "onh_kanitRaporu_metin", "onh_dosyaSurumu", "onh_firmaUnvan",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    # Motor I kolonları
    motor_map = {
        "onh_acikAdet": 38, "onh_acikTutar": 39, "onh_gecikenAdet": 40,
        "onh_gecikmeMaliyeti": 41, "onh_yetimKayit": 42, "onh_gecersizGecis": 43,
        "onh_zincirKirik": 44, "onh_yasKova": 45, "onh_tahminAralik": 46,
        "onh_senaryoKarsilastirma": 47, "onh_tornadoZirve": 48, "onh_oncelikSkoru": 49,
        "onh_paretoGecikme": 50, "onh_kapaliDusum": 51, "onh_kuyrukOzeti": 52,
        "onh_kanitRaporu": 53, "onh_sonrakiKimlik": 54,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "onh_yorumYas": 57, "onh_yorumGecikme": 58, "onh_yorumTahmin": 59,
        "onh_yorumSenaryo": 60, "onh_yorumTornado": 61, "onh_yorumOncelik": 62,
        "onh_yorumPareto": 63,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$13")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeOncelik", "LISTELER!$I$7:$I$9")


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
    dosya_ad = "OnUcHaftalikNakitOdemeOnceligi.xlsx"
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
