#!/usr/bin/env python3
"""İSG Risk Değerlendirme Pro (6331) — A4 üretim betiği (manda v6)."""

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
    yas_kova_formul,
    yaslandirma_gun,
    yetim_kayit_formul,
    zincir_kirik_formul,
)
from ortak.kanit_uretici import (  # noqa: E402
    imza_alani,
    rapor_baslik_satirlari,
    sha_placeholder_notu,
)

URUN_AD = "İSG Risk Değerlendirme Pro (6331)"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 200
DEMO_TEHLIKE = 24
DEMO_EYLEM = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 10)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "PLANLANDI", "durum_adi": "Planlandı", "sira": 1,
     "izinli_sonraki_durumlar": "DEVAM|IPTAL", "kategori": "acik"},
    {"durum_id": "DEVAM", "durum_adi": "Devam", "sira": 2,
     "izinli_sonraki_durumlar": "BEKLIYOR|TAMAMLANDI|IPTAL", "kategori": "acik"},
    {"durum_id": "BEKLIYOR", "durum_adi": "Bekliyor", "sira": 3,
     "izinli_sonraki_durumlar": "DEVAM|TAMAMLANDI|IPTAL", "kategori": "acik"},
    {"durum_id": "TAMAMLANDI", "durum_adi": "Tamamlandı", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "PLANLANDI|DEVAM",
    "PLANLANDI|IPTAL",
    "DEVAM|BEKLIYOR",
    "DEVAM|TAMAMLANDI",
    "DEVAM|IPTAL",
    "BEKLIYOR|DEVAM",
    "BEKLIYOR|TAMAMLANDI",
    "BEKLIYOR|IPTAL",
]

TEHLIKELER = [
    ("TEH-00001", "Üretim", "Makine koruyucu eksik", "ISG-001", 4, 5, 3, "Kısmi koruyucu"),
    ("TEH-00002", "Üretim", "Gürültü eşiği aşımı", "ISG-002", 3, 4, 4, "Kulaklık"),
    ("TEH-00003", "Depo", "Yüksekte istifleme", "ISG-001", 3, 5, 2, "Emniyet kemeri"),
    ("TEH-00004", "Depo", "Forklift çarpışma riski", "ISG-001", 4, 5, 3, "Yaya yolu"),
    ("TEH-00005", "Ofis", "Ergonomi / VDU", "ISG-003", 2, 3, 5, "Ayarlı masa"),
    ("TEH-00006", "Ofis", "Elektrik kablosu dağınıklığı", "ISG-001", 2, 4, 2, "Kabinet"),
    ("TEH-00007", "Saha", "Kaygan zemin", "ISG-001", 3, 4, 3, "Uyarı levhası"),
    ("TEH-00008", "Saha", "Kimyasal dökülme", "ISG-002", 3, 5, 2, "Spill kiti"),
    ("TEH-00009", "Üretim", "Toz maruziyeti", "ISG-002", 4, 4, 4, "Maske"),
    ("TEH-00010", "Üretim", "Sıcak yüzey teması", "ISG-001", 3, 4, 3, "Eldiven"),
    ("TEH-00011", "Bakım", "Enerji kesilmeden müdahale", "ISG-001", 5, 5, 2, "LOTO kısmi"),
    ("TEH-00012", "Bakım", "İskele stabilite", "ISG-001", 3, 5, 2, "Kontrol formu"),
    ("TEH-00013", "Lojistik", "Elle kaldırma", "ISG-003", 3, 3, 4, "Eğitim"),
    ("TEH-00014", "Lojistik", "Araç manevra alanı dar", "ISG-001", 4, 4, 3, "Ayna"),
    ("TEH-00015", "Mutfak", "Yağ yangını riski", "ISG-001", 2, 5, 2, "Yangın tüpü"),
    ("TEH-00016", "Mutfak", "Kesici alet yaralanma", "ISG-001", 3, 3, 4, "Eldiven"),
    ("TEH-00017", "Laboratuvar", "Solvent buharı", "ISG-002", 3, 4, 3, "Çeker ocak"),
    ("TEH-00018", "Laboratuvar", "Cam kırığı", "ISG-001", 2, 3, 3, "Kutu"),
    ("TEH-00019", "İnşaat", "Düşme yüksekliği", "ISG-001", 4, 5, 3, "Korkuluk"),
    ("TEH-00020", "İnşaat", "Ağır ekipman çarpma", "ISG-001", 3, 5, 2, "Bariyer"),
    ("TEH-00021", "İdari", "Acil çıkış engeli", "ISG-001", 2, 5, 1, "İşaret"),
    ("TEH-00022", "İdari", "Yangın tatbikatı eksik", "ISG-003", 2, 4, 1, "Plan"),
    ("TEH-00023", "Üretim", "Basınçlı hava hortumu", "ISG-001", 3, 3, 3, "Kelepçe"),
    ("TEH-00024", "Saha", "Gece aydınlatma yetersiz", "ISG-001", 3, 3, 2, "Projektör"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Risk skoru ve eylem planını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_eylem():
    satirlar = []
    durumlar = ["PLANLANDI", "DEVAM", "BEKLIYOR", "TAMAMLANDI", "IPTAL",
                "DEVAM", "BEKLIYOR", "PLANLANDI"]
    oncekiler = ["", "PLANLANDI", "DEVAM", "BEKLIYOR", "DEVAM",
                 "PLANLANDI", "DEVAM", "PLANLANDI"]
    for i in range(1, DEMO_EYLEM + 1):
        teh = TEHLIKELER[(i - 1) % DEMO_TEHLIKE][0]
        if i in (22, 25):
            teh = "" if i == 22 else "TEH-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "PLANLANDI", "TAMAMLANDI"
        if i == 20:
            o, d = "TAMAMLANDI", "DEVAM"
        o_, s_, _f = TEHLIKELER[(i - 1) % DEMO_TEHLIKE][4:7]
        ilk = o_ * s_
        hedef = max(1, ilk - (i % 5) * 2)
        if i in (12, 19):
            hedef = ilk + 40
        tar = RAPOR_TARIHI - timedelta(days=(i * 4) % 90)
        termin = RAPOR_TARIHI + timedelta(days=(i % 20) - 10)
        satirlar.append({
            "id": f"EYL-{i:05d}",
            "tehlike": teh,
            "tarih": tar,
            "termin": termin,
            "onceki": o,
            "durum": d,
            "sorumlu": f"Sorumlu {(i % 6) + 1}",
            "aciklama": f"Kontrol eylemi {i}",
            "ilk": ilk,
            "hedef": hedef,
            "tutar": max(1000, ilk * 250),
        })
    return satirlar


ORNEK_EYLEM = _demo_eylem()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "6331 sayılı Kanun çerçevesinde tehlike envanteri, L-matris / Fine-Kinney "
      "puanlama, eylem durum makinesi ve denetime sunulabilir kanıt raporu üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz eylem durum makinesi: Planlandı → Devam → Bekliyor → Tamamlandı / İptal",
        "L-matris veya Fine-Kinney seçimli puanlama (AYARLAR)",
        "Eylem yaşlandırma kovaları ve geçersiz geçiş uyarısı",
        "Skor dağılımı, uyum oranı, yenileme takvimi",
        "KANIT_RAPORU — madde atfı, revizyon, imza alanı",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "İSG uzmanı — envanter ve skorlama",
        "İşveren vekili — eylem onayı",
        "Denetçi — kanıt dosyası incelemesi",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) ENVANTER'e tehlikeleri girin. 2) AYARLAR'da matris tipini seçin. "
      "3) EYLEM_PLANI'nda kontrolleri ilerletin. 4) PANO ve KANIT_RAPORU'nu inceleyin.",
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
      '=IF(COUNTA(tblEnvanter[TehlikeId])=0,"VERİ YOK",'
      'IF(OR(isg_yetimKayit>0,isg_gecersizGecis>0,isg_zincirKirik>0,'
      'isg_kritikAdet>isg_esikKritik,isg_gecikenAdet>isg_esikGeciken),"İNCELE","UYGUN"))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Tehlike Adet", "=isg_tehlikeAdet", CATI),
        (7, "Kritik Adet", "=isg_kritikAdet", CATI),
        (8, "Ortalama Skor", "=isg_ortalamaSkor", CATI),
        (9, "Uyum Oranı", "=isg_uyumOrani", YÜZDE),
        (10, "Açık Eylem", "=isg_acikAdet", CATI),
        (11, "Geciken Eylem", "=isg_gecikenAdet", CATI),
        (12, "Yetim Kayıt", "=isg_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=isg_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=isg_zincirKirik", CATI),
        (15, "Yaş Kova", "=isg_eylemYas", None),
        (16, "Tahmin Aralık", "=isg_tahminAralik", CATI),
        (17, "Senaryo Farkı", "=isg_senaryoKarsilastirma", CATI),
        (18, "Tornado Zirve", "=isg_tornadoZirve", CATI),
        (19, "Eylem Öncelik", "=isg_eylemOncelik", TL),
        (20, "Pareto Kritik", "=isg_paretoKritik", YÜZDE),
        (21, "Gecikme Maliyeti", "=isg_gecikmeMaliyeti", TL),
        (22, "Anomali Skor", "=isg_anomaliSayisi", CATI),
        (23, "Skor Dağılım", "=isg_skorDagilim", None),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=isg_yorumEylemYas"),
        (8, "=isg_yorumGecikme"),
        (9, "=isg_yorumTahmin"),
        (10, "=isg_yorumSenaryo"),
        (11, "=isg_yorumTornado"),
        (12, "=isg_yorumOncelik"),
        (13, "=isg_yorumPareto"),
        (14, "=isg_yorumAnomali"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Seviye")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Düşük", 6), ("Orta", 8), ("Yüksek", 6), ("Kritik", 4),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Yaş Kova")
    h(ws, 26, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 7), ("8-30", 9), ("31-60", 6), ("60+", 6),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Hafta")
    h(ws, 26, 8, "Yeni Tehlike")
    for i in range(8):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, 2 + (i % 4), sayi=CATI)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Skor")
    for i, (ad, t) in enumerate([
        ("Temkinli", 420), ("Baz", 510), ("İyimser", 380),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=CATI)

    h(ws, 26, 13, "Uyarı")
    h(ws, 26, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Geciken", 8),
    ], 27):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 26, 16, "Durum")
    h(ws, 26, 17, "Adet")
    for i, (ad, adet) in enumerate([
        ("P-Adet", 5), ("D-Adet", 8), ("B-Adet", 6),
        ("T-Adet", 5), ("I-Adet", 4),
    ], 27):
        h(ws, i, 16, ad)
        h(ws, i, 17, adet, sayi=CATI)

    # A05 — PANO'da canlı kuyruk COUNTIFS/SUMIFS izi
    h(ws, 33, 1, "Açık Kuyruk (COUNTIFS)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 34, 1, "Planlandı kuyruk")
    h(ws, 34, 2, '=COUNTIFS(tblEylem[Durum],"PLANLANDI",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 34, 3, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"PLANLANDI")', sayi=TL)
    h(ws, 35, 1, "Devam kuyruk")
    h(ws, 35, 2, '=COUNTIFS(tblEylem[Durum],"DEVAM",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 35, 3, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"DEVAM")', sayi=TL)
    h(ws, 36, 1, "Bekliyor kuyruk")
    h(ws, 36, 2, '=COUNTIFS(tblEylem[Durum],"BEKLIYOR",tblEylem[Tutar],">0")', sayi=CATI)
    h(ws, 36, 3, '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"BEKLIYOR")', sayi=TL)

    c1 = PieChart()
    c1.title = "Skor Seviye Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=26, max_row=30), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=30))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A38")

    c2 = BarChart()
    c2.title = "Eylem Yaşlandırma"
    c2.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F38")

    c3 = LineChart()
    c3.title = "Haftalık Yeni Tehlike"
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

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=26, max_row=30), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=27, max_row=30))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A68")

    c6 = PieChart()
    c6.title = "Eylem Durum Payı"
    c6.add_data(Reference(ws, min_col=17, min_row=26, max_row=31), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=27, max_row=31))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F68")

    c7 = BarChart()
    c7.title = "Durum Adet"
    c7.add_data(Reference(ws, min_col=17, min_row=26, max_row=31), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=16, min_row=27, max_row=31))
    c7.width, c7.height = 10, 7
    ws.add_chart(c7, "A83")

    c8 = BarChart()
    c8.title = "Kova Payı"
    c8.add_data(Reference(ws, min_col=5, min_row=26, max_row=30), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=4, min_row=27, max_row=30))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F83")

    baski_hazirla(ws, "A1:N36", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def envanter(ws):
    sayfa_hazirla(ws, "ENVANTER", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Tehlike envanteri — her kaydın tek kimliği vardır (ID disiplini)",
             son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=isg_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni satırda kullanın | Örnek: TEH-00025 | Risk: Elle çakışma")

    sutunlar = [
        "TehlikeId", "Bolum", "TehlikeAdi", "MaddeAtif", "Olasilik", "Siddet",
        "Frekans", "MevcutKontrol", "Aktif", "HamSkor", "RiskSeviye", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "HamSkor": (
            'IF(tblEnvanter[[#This Row],[TehlikeId]]="","",'
            'IF(isg_matrisTipi="FK",'
            'tblEnvanter[[#This Row],[Olasilik]]*tblEnvanter[[#This Row],[Frekans]]*'
            'tblEnvanter[[#This Row],[Siddet]],'
            'tblEnvanter[[#This Row],[Olasilik]]*tblEnvanter[[#This Row],[Siddet]]))'
        ),
        "RiskSeviye": (
            'IF(tblEnvanter[[#This Row],[TehlikeId]]="","",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=isg_esikKritikSeviye,"Kritik",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=isg_esikYuksek,"Yüksek",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=isg_esikOrta,"Orta","Düşük"))))'
        ),
        "KayitDolu": 'IF(tblEnvanter[[#This Row],[TehlikeId]]="","",1)',
    }
    for i, (tid, bolum, ad, madde, ol, sid, fr, kontrol) in enumerate(TEHLIKELER):
        r = ILK + i
        _sari(ws, r, 1, tid, baslik="Tehlike Kimliği", mesaj="TEH-##### formatında")
        _sari(ws, r, 2, bolum, baslik="Bölüm", mesaj="İş alanı / bölüm")
        _sari(ws, r, 3, ad, baslik="Tehlike Adı", mesaj="Kısa tehlike tanımı")
        _sari(ws, r, 4, madde, baslik="Madde Atıf", mesaj="ISG-001 / ISG-002 / ISG-003")
        _sari(ws, r, 5, ol, sayi=CATI, baslik="Olasılık", mesaj="1-5 arası")
        _sari(ws, r, 6, sid, sayi=CATI, baslik="Şiddet", mesaj="1-5 arası")
        _sari(ws, r, 7, fr, sayi=CATI, baslik="Frekans", mesaj="Fine-Kinney için 1-5")
        _sari(ws, r, 8, kontrol, baslik="Mevcut Kontrol", mesaj="Mevcut önlem özeti")
        _sari(ws, r, 9, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_TEHLIKE, SON + 1):
        for c in range(1, 10):
            fmt = CATI if c in (5, 6, 7) else None
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblEnvanter", f"A{HDR}:L{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Tehlike Kimliği", mesaj="TEH-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Bölüm"), (3, "Tehlike"), (4, "Madde"), (8, "Kontrol")]:
        dogrulama(ws, "textLength", "0",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(5, "Olasılık"), (6, "Şiddet"), (7, "Frekans")]:
        dogrulama(ws, "whole", "1",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="1 ile 5 arası",
                  hata_baslik=ad, hata_mesaj="1-5 arası girin",
                  isaret="between", f2="5")
    dogrulama(ws, "list", "ListeEvetHayir", f"I{ILK}:I{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeMadde", f"D{ILK}:D{SON}",
              baslik="Madde Atıf", mesaj="ISG madde kodu seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["H"].width = 18


def degerlendirme(ws):
    sayfa_hazirla(ws, "DEGERLENDIRME", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; puanlama matrisi AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Tehlike toplam", '=COUNTA(tblEnvanter[TehlikeId])'),
        ("H02", "Dolu kayıt", '=SUM(tblEnvanter[KayitDolu])'),
        ("H03", "Aktif tehlike", '=COUNTIF(tblEnvanter[Aktif],"EVET")'),
        ("H04", "Skor toplam", "=SUM(tblEnvanter[HamSkor])"),
        ("H05", "Ortalama skor", "=IFERROR(AVERAGE(tblEnvanter[HamSkor]),0)"),
        ("H06", "Maks skor", "=IFERROR(MAX(tblEnvanter[HamSkor]),0)"),
        ("H07", "Min skor", "=IFERROR(MIN(tblEnvanter[HamSkor]),0)"),
        ("H08", "Düşük adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Düşük")'),
        ("H09", "Orta adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Orta")'),
        ("H10", "Yüksek adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Yüksek")'),
        ("H11", "Kritik adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Kritik")'),
        ("H12", "ISG-001 atıf", '=COUNTIF(tblEnvanter[MaddeAtif],"ISG-001")'),
        ("H13", "ISG-002 atıf", '=COUNTIF(tblEnvanter[MaddeAtif],"ISG-002")'),
        ("H14", "ISG-003 atıf", '=COUNTIF(tblEnvanter[MaddeAtif],"ISG-003")'),
        ("H15", "Eylem toplam", '=COUNTA(tblEylem[EylemId])'),
        ("H16", "Planlandı", '=COUNTIF(tblEylem[Durum],"PLANLANDI")'),
        ("H17", "Devam", '=COUNTIF(tblEylem[Durum],"DEVAM")'),
        ("H18", "Bekliyor", '=COUNTIF(tblEylem[Durum],"BEKLIYOR")'),
        ("H19", "Tamamlandı", '=COUNTIF(tblEylem[Durum],"TAMAMLANDI")'),
        ("H20", "İptal", '=COUNTIF(tblEylem[Durum],"IPTAL")'),
        ("H21", "Açık adet", "=C21+C22+C23"),
        ("H22", "Kapalı adet", "=C24+C25"),
        ("H23", "Açık tutar",
         '=SUMIF(tblEylem[Durum],"PLANLANDI",tblEylem[Tutar])'
         '+SUMIF(tblEylem[Durum],"DEVAM",tblEylem[Tutar])'
         '+SUMIF(tblEylem[Durum],"BEKLIYOR",tblEylem[Tutar])'),
        ("H24", "Yetim adet", '=COUNTIF(tblEylem[YetimBayrak],"YETİM")'),
        ("H25", "Geçersiz geçiş", '=COUNTIF(tblEylem[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H26", "Zincir kırık", '=COUNTIF(tblEylem[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H27", "Zincir OK", '=COUNTIF(tblEylem[ZincirBayrak],"OK")'),
        ("H28", "Yaş 0-7", '=COUNTIF(tblEylem[YasKova],"0-7")'),
        ("H29", "Yaş 8-30", '=COUNTIF(tblEylem[YasKova],"8-30")'),
        ("H30", "Yaş 31-60", '=COUNTIF(tblEylem[YasKova],"31-60")'),
        ("H31", "Yaş 60+", '=COUNTIF(tblEylem[YasKova],"60+")'),
        ("H32", "Ort yaş gün", "=IFERROR(AVERAGE(tblEylem[YasGun]),0)"),
        ("H33", "Geciken adet",
         '=COUNTIFS(tblEylem[YasGun],">"&isg_esikGecikmeGun,tblEylem[Durum],"<>TAMAMLANDI",'
         'tblEylem[Durum],"<>IPTAL",tblEylem[Tutar],">0")'),
        ("H34", "Geciken tutar",
         '=SUMIFS(tblEylem[Tutar],tblEylem[YasGun],">"&isg_esikGecikmeGun,'
         'tblEylem[Durum],"<>TAMAMLANDI",tblEylem[Durum],"<>IPTAL")'),
        ("H35", "Gecikme maliyeti", "=C39*isg_gecikmeOranGun"),
        ("H36", "Temkinli senaryo", "=C9*isg_senaryoTemkinli"),
        ("H37", "Baz senaryo", "=C9*isg_senaryoBaz"),
        ("H38", "İyimser senaryo", "=C9*isg_senaryoIyimser"),
        ("H39", "Senaryo farkı", "=C41-C43"),
        ("H40", "Tornado etki", "=C9*isg_tornadoOran"),
        ("H41", "Eylem öncelik havuzu",
         '=SUMIFS(tblEylem[Tutar],tblEylem[Durum],"BEKLIYOR",'
         'tblEylem[YasGun],">"&isg_esikGecikmeGun)'),
        ("H42", "Pareto kritik payı", "=IFERROR(C16/MAX(C6,1),0)"),
        ("H43", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblEnvanter[HamSkor],isg_yuzdelikOran),0)"),
        ("H44", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblEnvanter[TehlikeId])+1,tblEnvanter[HamSkor],"
         "tblEnvanter[KayitDolu]),C48)"),
        ("H45", "Tahmin aralık", "=IFERROR((C48+C49)/2,0)"),
        ("H46", "Uyum oranı", "=IFERROR(C24/MAX(C20,1),0)"),
        ("H47", "Anomali skor",
         '=COUNTIFS(tblEnvanter[HamSkor],">"&isg_esikAnomali)+C29+C30+C31'),
        ("H48", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C33,C34,C35,C36)'),
        ("H49", "Skor dağılım metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"D",C13,"O",C14,"Y",C15,"K",C16)'),
        ("H50", "Kuyruk özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Plan",C21,"Devam",C22,"Bekliyor",C23,"Geciken",C38)'),
        ("H51", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C29*10-C30*10-C31*10-C16*5-C38*2))"),
        ("H52", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H53", "Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"ISG-001",C17,"ISG-002",C18,"ISG-003",C19)'),
        ("H54", "Yenileme kalan gün", "=isg_yenilemeTarihi-raporTarihi"),
        ("H55", "Kapalı düşüm", "=C27"),
    ]
    for i, (kod, acik, form) in enumerate(adimlar):
        r = 6 + i
        h(ws, r, 1, kod)
        h(ws, r, 2, acik)
        fmt = TL if any(x in acik.lower() for x in (
            "tutar", "maliyet", "öncelik", "havuz")) and "oran" not in acik.lower() else (
            YÜZDE if "oran" in acik.lower() or "payı" in acik.lower() or "hazırlık" in acik.lower()
            else (GUN if "yaş gün" in acik.lower() or "kalan gün" in acik.lower()
                  else (None if "metin" in acik.lower() or "özet" in acik.lower()
                        or "dağılım" in acik.lower() else CATI))
        )
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 64, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 64, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 64, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 65, 1, "Temkinli")
    h(ws, 65, 2, "=isg_senaryoTemkinli*IF(COUNTA(tblEnvanter[TehlikeId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 65, 3, "=C41", sayi=CATI)
    h(ws, 66, 1, "Baz")
    h(ws, 66, 2, "=isg_senaryoBaz*IF(COUNTA(tblEnvanter[TehlikeId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 66, 3, "=C42", sayi=CATI)
    h(ws, 67, 1, "İyimser")
    h(ws, 67, 2, "=isg_senaryoIyimser*IF(COUNTA(tblEnvanter[TehlikeId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 67, 3, "=C43", sayi=CATI)

    h(ws, 69, 1, "Kuyruk (açık eylemler)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 70, 1, "Planlandı")
    h(ws, 70, 2, kuyruk_sayim("tblEylem", "PLANLANDI"), sayi=CATI)
    h(ws, 70, 3, kuyruk_tutar("tblEylem", "PLANLANDI"), sayi=TL)
    h(ws, 71, 1, "Devam")
    h(ws, 71, 2, kuyruk_sayim("tblEylem", "DEVAM"), sayi=CATI)
    h(ws, 71, 3, kuyruk_tutar("tblEylem", "DEVAM"), sayi=TL)
    h(ws, 72, 1, "Bekliyor")
    h(ws, 72, 2, kuyruk_sayim("tblEylem", "BEKLIYOR"), sayi=CATI)
    h(ws, 72, 3, kuyruk_tutar("tblEylem", "BEKLIYOR"), sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 55})


def eylem_plani(ws):
    sayfa_hazirla(ws, "EYLEM_PLANI", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Eylem satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "EylemId", "KaynakTehlikeId", "Tarih", "Termin", "OncekiDurum", "Durum",
        "Sorumlu", "Aciklama", "IlkSkor", "HedefSkor", "Tutar",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblEylem[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblEylem[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblEylem[[#This Row],[KaynakTehlikeId]]",
        "tblEnvanter[TehlikeId]").lstrip("=")
    gecis_birlesik = (
        'tblEylem[[#This Row],[OncekiDurum]]&"|"&tblEylem[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblEylem[[#This Row],[IlkSkor]]",
        "tblEylem[[#This Row],[HedefSkor]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "YasGun": f'IF(tblEylem[[#This Row],[EylemId]]="","",{yas_f})',
        "YasKova": f'IF(tblEylem[[#This Row],[EylemId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblEylem[[#This Row],[EylemId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblEylem[[#This Row],[EylemId]]="",'
            f'tblEylem[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblEylem[[#This Row],[EylemId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblEylem[[#This Row],[EylemId]]="","",1)',
    }

    for i, s in enumerate(ORNEK_EYLEM):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Eylem Kimliği", mesaj="EYL-#####")
        _sari(ws, r, 2, s["tehlike"] or None, baslik="Kaynak Tehlike",
              mesaj="ENVANTER TehlikeId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Açılış tarihi")
        _sari(ws, r, 4, s["termin"], sayi=TARİH, baslik="Termin", mesaj="Hedef bitiş")
        _sari(ws, r, 5, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 7, s["sorumlu"], baslik="Sorumlu", mesaj="Sorumlu kişi")
        _sari(ws, r, 8, s["aciklama"], baslik="Açıklama", mesaj="Kısa eylem açıklaması")
        _sari(ws, r, 9, s["ilk"], sayi=CATI, baslik="İlk Skor", mesaj="Başlangıç risk skoru")
        _sari(ws, r, 10, s["hedef"], sayi=CATI, baslik="Hedef Skor", mesaj="Hedef kalan risk")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Tahmini maliyet TL")

    for r in range(ILK + DEMO_EYLEM, SON + 1):
        for c in range(1, 12):
            fmt = TARİH if c in (3, 4) else (TL if c == 11 else (CATI if c in (9, 10) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblEylem", f"A{HDR}:Q{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Eylem Kimliği", mesaj="EYL-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Tehlike", mesaj="Tehlike kimliği (boş = yetim riski)",
              hata_baslik="Tehlike", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="Açılış tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "date", "1", f"D{ILK}:D{SON}",
              baslik="Termin", mesaj="Hedef bitiş tarihi",
              hata_baslik="Termin", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    dogrulama(ws, "list", "ListeDurumlar", f"E{ILK}:E{SON}",
              baslik="Önceki Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeDurumlar", f"F{ILK}:F{SON}",
              baslik="Durum", mesaj="Durum haritasından seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(9, "İlk Skor"), (10, "Hedef Skor")]:
        dogrulama(ws, "whole", "0",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"K{ILK}:K{SON}",
              baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    for c, ad in [(7, "Sorumlu"), (8, "Açıklama")]:
        dogrulama(ws, "textLength", "0",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 18)})
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["O"].width = 16


def akis_ozet(ws):
    """A02/A05 için AKIS sayfası — EYLEM_PLANI ile aynı motor, yetim/kuyruk izi."""
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3,
             "Eylem akış özeti — KAYNAK_KART_ID / yetim / kuyruk (EYLEM_PLANI ile bağlı)",
             son_kolon=10)
    h(ws, 5, 1, "KAYNAK_KART_ID kontrolü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Yetim (YETİM) adet")
    h(ws, 6, 2, '=COUNTIF(tblEylem[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 7, 1, "Geçersiz geçiş")
    h(ws, 7, 2, '=COUNTIF(tblEylem[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 8, 1, "Zincir kırık")
    h(ws, 8, 2, '=COUNTIF(tblEylem[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)
    h(ws, 10, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 11, 1, "Planlandı")
    h(ws, 11, 2, kuyruk_sayim("tblEylem", "PLANLANDI"), sayi=CATI)
    h(ws, 11, 3, kuyruk_tutar("tblEylem", "PLANLANDI"), sayi=TL)
    h(ws, 12, 1, "Devam")
    h(ws, 12, 2, kuyruk_sayim("tblEylem", "DEVAM"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblEylem", "DEVAM"), sayi=TL)
    h(ws, 13, 1, "Bekliyor")
    h(ws, 13, 2, kuyruk_sayim("tblEylem", "BEKLIYOR"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblEylem", "BEKLIYOR"), sayi=TL)
    h(ws, 15, 1, "Kapalı (Tamamlandı+İptal) kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Tamamlandı")
    h(ws, 16, 2, kuyruk_sayim("tblEylem", "TAMAMLANDI"), sayi=CATI)
    h(ws, 17, 1, "İptal")
    h(ws, 17, 2, kuyruk_sayim("tblEylem", "IPTAL"), sayi=CATI)
    genislik(ws, {"A": 42, "B": 14, "C": 16})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Denetim Kanıt Raporu — 6331 Risk Değerlendirme", boyut=16, son_kolon=10)
    for i, (etiket, deger) in enumerate(rapor_baslik_satirlari(URUN_AD, SURUM), 4):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger)
    h(ws, 8, 1, "Rapor Tarihi")
    h(ws, 8, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 9, 1, "KARAR", kalin=True)
    h(ws, 9, 2, "=PANO!B4", kalin=True, boyut=14)

    ozet = [
        (11, "Tehlike Adet", "=isg_tehlikeAdet", CATI),
        (12, "Kritik Adet", "=isg_kritikAdet", CATI),
        (13, "Uyum Oranı", "=isg_uyumOrani", YÜZDE),
        (14, "Yetim", "=isg_yetimKayit", CATI),
        (15, "Geçersiz Geçiş", "=isg_gecersizGecis", CATI),
        (16, "Zincir Kırık", "=isg_zincirKirik", CATI),
        (17, "Madde Atıf", "=isg_maddeAtif", None),
        (18, "Kanıt Özeti", "=isg_kanitRaporu", None),
        (19, "Skor Dağılım", "=isg_skorDagilim", None),
        (20, "Yenileme Takvim", "=isg_yenilemeTakvim", TARİH),
        (21, "Revizyon Geçmişi", "=isg_revizyonGecmisi", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)

    h(ws, 23, 1, "Hesap Zinciri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1,
      "Olasılık × Şiddet (L) veya Olasılık × Frekans × Şiddet (FK) → HamSkor → "
      "RiskSeviye → Eylem durum makinesi → PANO kararı.",
      kaydir=True)
    ws.merge_cells("A24:H24")

    h(ws, 26, 1, "Varsayımlar ve Belirsizlik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1,
      "Matris seçimi (L-matris / Fine-Kinney) kullanıcı tercihidir; aynı girdiler "
      "farklı skor üretebilir. Bu çıktı karar destektir; kesin hukuki/iş güvenliği "
      "görüşü yerine geçmez. KVKK: kişisel veri toplanmaz, çevrimdışı çalışır.",
      kaydir=True)
    ws.merge_cells("A27:H27")
    h(ws, 28, 1, "=isg_kvkkBeyani", kaydir=True)
    ws.merge_cells("A28:H28")

    h(ws, 30, 1, "İmza / Onay", kalin=True, yazi=KOYU_LACIVERT)
    for i, (etiket, deger) in enumerate(imza_alani(), 31):
        h(ws, i, 1, etiket)
        if deger:
            h(ws, i, 2, deger, sayi=TARİH if "Tarih" in etiket else None)
        else:
            _sari(ws, i, 2, "", baslik=etiket, mesaj="İmza / unvan yazın")

    h(ws, 36, 1, sha_placeholder_notu(), kaydir=True, yazi=GRİ, boyut=9)
    ws.merge_cells("A36:H36")
    h(ws, 38, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H38", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "Tehlike sayısı", '=COUNTA(tblEnvanter[TehlikeId])'),
        (6, "Eylem sayısı", '=COUNTA(tblEylem[EylemId])'),
        (7, "Kritik", "=isg_kritikAdet"),
        (8, "Planlandı", '=COUNTIF(tblEylem[Durum],"PLANLANDI")'),
        (9, "Devam", '=COUNTIF(tblEylem[Durum],"DEVAM")'),
        (10, "Bekliyor", '=COUNTIF(tblEylem[Durum],"BEKLIYOR")'),
        (11, "Tamamlandı", '=COUNTIF(tblEylem[Durum],"TAMAMLANDI")'),
        (12, "Yetim", "=isg_yetimKayit"),
        (13, "Geçersiz", "=isg_gecersizGecis"),
        (14, "Zincir kırık", "=isg_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=DEGERLENDIRME!C56"),
        (17, "Uyum oranı", "=isg_uyumOrani"),
        (18, "Gecikme maliyeti", "=isg_gecikmeMaliyeti"),
        (19, "Kapalı düşüm", "=isg_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        fmt = YÜZDE if r == 17 else (TL if r == 18 else (CATI if r < 15 else None))
        h(ws, r, 2, form, kalin=True, sayi=fmt)
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_TEHLIKE} tehlike ve {DEMO_EYLEM} eylem satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "ENVANTER ve EYLEM_PLANI'nı temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Planlandı/Devam/Bekliyor/Tamamlandı/İptal — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 5000, sayi=CATI)
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

    h(ws, 3, 7, "Madde Atıf", kalin=True)
    h(ws, 6, 7, "Madde", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, m in enumerate(["ISG-001", "ISG-002", "ISG-003"], 7):
        _sari(ws, i, 7, m, baslik="Madde", mesaj="6331 madde kodu")

    h(ws, 3, 9, "Matris Tipi", kalin=True)
    h(ws, 6, 9, "Tip", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 9, "L", baslik="Matris", mesaj="L-matris")
    _sari(ws, 8, 9, "FK", baslik="Matris", mesaj="Fine-Kinney")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 12, "I": 10})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası + puanlama matrisi",
             son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 15, "skor", "İş kuralı — skor sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği (ilk-hedef)"),
        ("isg_matrisTipi", "L", "metin", "Kullanıcı seçimi L/FK", "01.01.2026",
         "L-matris veya Fine-Kinney"),
        ("isg_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("isg_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("isg_esikKritik", 3, "adet", "İç politika", "01.01.2026", "Karar eşiği kritik"),
        ("isg_esikKritikSeviye", 20, "skor", "L-matris eşik", "01.01.2026", "Kritik skor eşiği"),
        ("isg_esikYuksek", 12, "skor", "L-matris eşik", "01.01.2026", "Yüksek skor eşiği"),
        ("isg_esikOrta", 6, "skor", "L-matris eşik", "01.01.2026", "Orta skor eşiği"),
        ("isg_esikAnomali", 25, "skor", "Anomali kuralı", "01.01.2026", "Anomali skor eşiği"),
        ("isg_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("isg_senaryoTemkinli", 1.15, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("isg_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("isg_senaryoIyimser", 0.85, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("isg_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("isg_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("isg_yenilemeTarihi", date(2027, 8, 10), "tarih", "Yıllık yenileme", "01.01.2026",
         "Sonraki risk değerlendirme"),
        ("isg_olcekHedef", 5000, "satir", "Manda A4 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("isg_azamiSkor", 500, "skor", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("isg_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("isg_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("isg_puanlamaMatrisi", "L veya FK AYARLAR", "metin", "SPEC puanlama", "01.01.2026",
         "Puanlama matrisi kaynağı"),
        ("isg_eylemDurum", "tblDurumHaritasi", "metin", "SPEC akis", "01.01.2026",
         "Eylem durum makinesi"),
        ("isg_kvkkBeyani",
         "Kişisel veri toplanmaz; dosya çevrimdışı ve makrosuzdur.",
         "metin", "KVKK beyanı", "01.01.2026", "KVKK / çevrimdışı beyan"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 2:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    dogrulama(ws, "list", "ListeMatris", "B8",
              baslik="Matris Tipi", mesaj="L veya FK seçin",
              hata_baslik="Liste", hata_mesaj="L veya FK")

    h(ws, 32, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 33, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 34 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H33:L38", dh_bas)

    h(ws, 41, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 42, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 42, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor = [
        (43, "isg_tehlikeAdet", "=DEGERLENDIRME!C6"),
        (44, "isg_kritikAdet", "=DEGERLENDIRME!C16"),
        (45, "isg_ortalamaSkor", "=DEGERLENDIRME!C10"),
        (46, "isg_uyumOrani", "=DEGERLENDIRME!C51"),
        (47, "isg_acikAdet", "=DEGERLENDIRME!C26"),
        (48, "isg_gecikenAdet", "=DEGERLENDIRME!C38"),
        (49, "isg_yetimKayit", "=DEGERLENDIRME!C29"),
        (50, "isg_gecersizGecis", "=DEGERLENDIRME!C30"),
        (51, "isg_zincirKirik", "=DEGERLENDIRME!C31"),
        (52, "isg_eylemYas", "=DEGERLENDIRME!C53"),
        (53, "isg_tahminAralik", "=DEGERLENDIRME!C50"),
        (54, "isg_senaryoKarsilastirma", "=DEGERLENDIRME!C44"),
        (55, "isg_tornadoZirve", "=DEGERLENDIRME!C45"),
        (56, "isg_eylemOncelik", "=DEGERLENDIRME!C46"),
        (57, "isg_paretoKritik", "=DEGERLENDIRME!C47"),
        (58, "isg_gecikmeMaliyeti", "=DEGERLENDIRME!C40"),
        (59, "isg_anomaliSayisi", "=DEGERLENDIRME!C52"),
        (60, "isg_skorDagilim", "=DEGERLENDIRME!C54"),
        (61, "isg_kapaliDusum", "=DEGERLENDIRME!C60"),
        (62, "isg_kuyrukOzeti", "=DEGERLENDIRME!C55"),
        (63, "isg_maddeAtif", "=DEGERLENDIRME!C58"),
        (64, "isg_yenilemeTakvim", "=isg_yenilemeTarihi"),
        (65, "isg_revizyonGecmisi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Sürüm",isg_dosyaSurumu,"Rapor",TEXT(raporTarihi,"dd.mm.yyyy"))'),
        (66, "isg_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Tehlike",DEGERLENDIRME!C6,"Kritik",DEGERLENDIRME!C16,'
         '"Yetim",DEGERLENDIRME!C29,"Uyum",TEXT(DEGERLENDIRME!C51,"0%"))'),
        (67, "isg_sonrakiKimlik",
         '="TEH-"&TEXT(COUNTA(tblEnvanter[TehlikeId])+1,"00000")'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 69, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (70, "isg_yorumEylemYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",isg_eylemYas)'),
        (71, "isg_yorumGecikme",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Gecikme maliyeti",TEXT(isg_gecikmeMaliyeti,"0.00"),"TL")'),
        (72, "isg_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(isg_tahminAralik,"0.00"))'),
        (73, "isg_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(isg_senaryoKarsilastirma,"0.00"))'),
        (74, "isg_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(isg_tornadoZirve,"0.00"))'),
        (75, "isg_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Eylem öncelik havuzu",TEXT(isg_eylemOncelik,"0.00"),"TL")'),
        (76, "isg_yorumPareto",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Pareto kritik payı",TEXT(isg_paretoKritik,"0.0%"))'),
        (77, "isg_yorumAnomali",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Anomali sayısı",isg_anomaliSayisi)'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 36, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. ENVANTER sayfasına tehlikeleri girin; sonraki kimliği kullanın.",
        "2. AYARLAR'da isg_matrisTipi alanından L (L-matris) veya FK (Fine-Kinney) seçin.",
        "3. EYLEM_PLANI'nda her kontrol eylemini yazın; durum kolonunu listeden seçin.",
        "4. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "5. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "6. Kapalı durumlar (Tamamlandı/İptal) açık kuyruklardan düşer.",
        "7. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "8. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "9. KANIT_RAPORU sayfasını PDF olarak yazdırabilirsiniz.",
        "10. BELİRSİZLİK — matris_secimi_L_vs_FK: L-matris ile Fine-Kinney aynı girdide "
        "farklı skor üretir; hangi matrisin seçildiği AYARLAR.isg_matrisTipi ile kayıtlıdır. "
        "Denetimde matrisi ve eşikleri açıkça belirtin.",
        f"11. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 30):
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
    for r in range(7, 15):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Geçiş", mesaj="eski|yeni",
                  hata_baslik="Geçiş", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="40")
    ws = wb["KANIT_RAPORU"]
    for r, ad in [(5, "Sürüm"), (31, "Hazırlayan"), (32, "Kontrol"), (33, "Tarih")]:
        dogrulama(ws, "textLength", "0", f"B{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KILAVUZ"]
    for r in range(5, 16):
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
    for r in range(6, 24):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["EYLEM_PLANI"]
    for col, val, fill, font in [
        ("N", '"YETİM"', kirmizi, kf),
        ("O", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("P", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("P", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("BEKLIYOR", sari), ("TAMAMLANDI", yesil), ("IPTAL", kirmizi)]:
        ws.conditional_formatting.add(
            f"F{ILK}:F{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["ENVANTER"]
    for seviye, fill in [("Kritik", kirmizi), ("Yüksek", sari), ("Düşük", yesil)]:
        ws.conditional_formatting.add(
            f"K{ILK}:K{SON}",
            CellIsRule(operator="equal", formula=[f'"{seviye}"'], fill=fill))
    ws.conditional_formatting.add(
        f"I{ILK}:I{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))

    ws = wb["DEGERLENDIRME"]
    for r in range(6, 61):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    for r in range(70, 73):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))

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

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(11, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "isg_matrisTipi", "isg_esikGecikmeGun",
        "isg_esikGeciken", "isg_esikKritik", "isg_esikKritikSeviye", "isg_esikYuksek",
        "isg_esikOrta", "isg_esikAnomali", "isg_gecikmeOranGun", "isg_senaryoTemkinli",
        "isg_senaryoBaz", "isg_senaryoIyimser", "isg_tornadoOran", "isg_yuzdelikOran",
        "isg_yenilemeTarihi", "isg_olcekHedef", "isg_azamiSkor", "isg_dosyaSurumu",
        "isg_firmaUnvan", "isg_puanlamaMatrisi", "isg_eylemDurum", "isg_kvkkBeyani",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "isg_tehlikeAdet": 43, "isg_kritikAdet": 44, "isg_ortalamaSkor": 45,
        "isg_uyumOrani": 46, "isg_acikAdet": 47, "isg_gecikenAdet": 48,
        "isg_yetimKayit": 49, "isg_gecersizGecis": 50, "isg_zincirKirik": 51,
        "isg_eylemYas": 52, "isg_tahminAralik": 53, "isg_senaryoKarsilastirma": 54,
        "isg_tornadoZirve": 55, "isg_eylemOncelik": 56, "isg_paretoKritik": 57,
        "isg_gecikmeMaliyeti": 58, "isg_anomaliSayisi": 59, "isg_skorDagilim": 60,
        "isg_kapaliDusum": 61, "isg_kuyrukOzeti": 62, "isg_maddeAtif": 63,
        "isg_yenilemeTakvim": 64, "isg_revizyonGecmisi": 65, "isg_kanitRaporu": 66,
        "isg_sonrakiKimlik": 67,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "isg_yorumEylemYas": 70, "isg_yorumGecikme": 71, "isg_yorumTahmin": 72,
        "isg_yorumSenaryo": 73, "isg_yorumTornado": 74, "isg_yorumOncelik": 75,
        "isg_yorumPareto": 76, "isg_yorumAnomali": 77,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$14")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeMadde", "LISTELER!$G$7:$G$9")
    ad_ekle(wb, "ListeMatris", "LISTELER!$I$7:$I$8")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (envanter, "ENVANTER"),
        (degerlendirme, "DEGERLENDIRME"),
        (eylem_plani, "EYLEM_PLANI"),
        (akis_ozet, "AKIS"),
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
    dosya_ad = "IsgRiskDegerlendirmePro6331.xlsx"
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
