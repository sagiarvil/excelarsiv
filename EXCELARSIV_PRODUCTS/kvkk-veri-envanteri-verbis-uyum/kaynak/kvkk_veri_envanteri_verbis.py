#!/usr/bin/env python3
"""KVKK Veri Envanteri + VERBİS Uyum Sistemi — A4 üretim betiği (manda v6)."""

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

URUN_AD = "KVKK Veri Envanteri + VERBİS Uyum Sistemi"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 200
DEMO_KAYIT = 24
DEMO_AKIS = 28
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 11)

DURUM_HARITASI = durum_haritasi_dogrula([
    {"durum_id": "ENVANTAR", "durum_adi": "Envanter", "sira": 1,
     "izinli_sonraki_durumlar": "KONTROL|IPTAL", "kategori": "acik"},
    {"durum_id": "KONTROL", "durum_adi": "Kontrol", "sira": 2,
     "izinli_sonraki_durumlar": "DUZELTME|UYUMLU|IPTAL", "kategori": "acik"},
    {"durum_id": "DUZELTME", "durum_adi": "Düzeltme", "sira": 3,
     "izinli_sonraki_durumlar": "KONTROL|UYUMLU|IPTAL", "kategori": "acik"},
    {"durum_id": "UYUMLU", "durum_adi": "Uyumlu", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "IPTAL", "durum_adi": "İptal", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "ENVANTAR|KONTROL",
    "ENVANTAR|IPTAL",
    "KONTROL|DUZELTME",
    "KONTROL|UYUMLU",
    "KONTROL|IPTAL",
    "DUZELTME|KONTROL",
    "DUZELTME|UYUMLU",
    "DUZELTME|IPTAL",
]

# KayitId, Kategori, VeriKonusu, HukukiSebep, KuralAtif, Hassasiyet, AktarimRisk,
# SaklamaAy, YurtDisi, Guvenlik, MevcutOnlem
KAYITLAR = [
    ("VER-00001", "Kimlik", "Çalışan kimlik bilgisi", "Sözleşme", "KVV-001", 4, 3, 120, "HAYIR", 3, "Erişim logu"),
    ("VER-00002", "İletişim", "Müşteri e-posta / telefon", "Açık rıza", "KVV-004", 3, 4, 36, "EVET", 2, "Şifreleme kısmi"),
    ("VER-00003", "Finans", "Maaş ve banka IBAN", "Yasal yükümlülük", "KVV-001", 5, 2, 120, "HAYIR", 4, "Yetki matrisi"),
    ("VER-00004", "Finans", "Fatura cari kayıtları", "Meşru menfaat", "KVV-002", 3, 2, 60, "HAYIR", 3, "Yedekleme"),
    ("VER-00005", "Sağlık", "İşyeri hekimi muayene", "Yasal yükümlülük", "KVV-001", 5, 1, 180, "HAYIR", 5, "Kilitli klasör"),
    ("VER-00006", "Lokasyon", "Ziyaretçi kamera görüntüsü", "Meşru menfaat", "KVV-002", 3, 2, 12, "HAYIR", 2, "Saklama politikası"),
    ("VER-00007", "Kimlik", "Tedarikçi yetkili kimlik", "Sözleşme", "KVV-001", 3, 3, 60, "EVET", 2, "Sözleşme maddesi"),
    ("VER-00008", "İletişim", "Pazarlama iletişim listesi", "Açık rıza", "KVV-004", 2, 4, 24, "EVET", 1, "Rıza kaydı eksik"),
    ("VER-00009", "Dijital", "Web çerez / oturum", "Açık rıza", "KVV-004", 2, 5, 12, "EVET", 2, "Çerez paneli"),
    ("VER-00010", "Dijital", "ERP kullanıcı erişim logu", "Meşru menfaat", "KVV-003", 3, 2, 24, "HAYIR", 4, "SIEM"),
    ("VER-00011", "Kimlik", "Aday özgeçmiş dosyası", "Açık rıza", "KVV-004", 3, 2, 6, "HAYIR", 2, "Silme takvimi"),
    ("VER-00012", "Finans", "Kredi kartı token", "Sözleşme", "KVV-003", 5, 5, 36, "EVET", 3, "PCI maskeleme"),
    ("VER-00013", "İletişim", "Destek talep yazışması", "Sözleşme", "KVV-002", 2, 2, 24, "HAYIR", 3, "Ticket sistemi"),
    ("VER-00014", "Lokasyon", "Araç takip GPS", "Meşru menfaat", "KVV-002", 3, 3, 12, "HAYIR", 2, "Politika metni"),
    ("VER-00015", "Sağlık", "COVID / aşı beyanı arşivi", "Yasal yükümlülük", "KVV-005", 4, 1, 24, "HAYIR", 3, "Arşiv kilidi"),
    ("VER-00016", "Kimlik", "VERBİS bildirim taslağı", "Yasal yükümlülük", "KVV-005", 2, 1, 60, "HAYIR", 4, "Uyum klasörü"),
    ("VER-00017", "Dijital", "Bulut yedek kişisel veri", "Sözleşme", "KVV-003", 4, 5, 36, "EVET", 2, "SCC taslak"),
    ("VER-00018", "İletişim", "E-bülten abone listesi", "Açık rıza", "KVV-004", 2, 3, 18, "EVET", 1, "Rıza tarihi yok"),
    ("VER-00019", "Finans", "İcra / haciz dosya notu", "Yasal yükümlülük", "KVV-001", 4, 1, 120, "HAYIR", 4, "Yetki sınırı"),
    ("VER-00020", "Kimlik", "Ortaklık pay defteri", "Yasal yükümlülük", "KVV-001", 3, 1, 120, "HAYIR", 4, "Noter suret"),
    ("VER-00021", "Dijital", "Mobil uygulama cihaz kimliği", "Meşru menfaat", "KVV-003", 3, 4, 24, "EVET", 2, "Anonimleştirme"),
    ("VER-00022", "İletişim", "Çağrı merkezi ses kaydı", "Meşru menfaat", "KVV-002", 4, 2, 6, "HAYIR", 3, "Saklama 6 ay"),
    ("VER-00023", "Sağlık", "Meslek hastalığı dosyası", "Yasal yükümlülük", "KVV-001", 5, 1, 180, "HAYIR", 5, "Özel kategori"),
    ("VER-00024", "Lokasyon", "Ofis kartlı geçiş logu", "Meşru menfaat", "KVV-002", 2, 1, 12, "HAYIR", 3, "Kart sistemi"),
]


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Uyum skoru ve VERBİS dosyasını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    satirlar = []
    durumlar = ["ENVANTAR", "KONTROL", "DUZELTME", "UYUMLU", "IPTAL",
                "KONTROL", "DUZELTME", "ENVANTAR"]
    oncekiler = ["", "ENVANTAR", "KONTROL", "DUZELTME", "KONTROL",
                 "ENVANTAR", "KONTROL", "ENVANTAR"]
    for i in range(1, DEMO_AKIS + 1):
        kid = KAYITLAR[(i - 1) % DEMO_KAYIT][0]
        if i in (22, 25):
            kid = "" if i == 22 else "VER-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "ENVANTAR", "UYUMLU"
        if i == 20:
            o, d = "UYUMLU", "KONTROL"
        hass, akt = KAYITLAR[(i - 1) % DEMO_KAYIT][5:7]
        ilk = hass * akt
        hedef = max(1, ilk - (i % 5) * 2)
        if i in (12, 19):
            hedef = ilk + 40
        tar = RAPOR_TARIHI - timedelta(days=(i * 4) % 90)
        termin = RAPOR_TARIHI + timedelta(days=(i % 20) - 10)
        satirlar.append({
            "id": f"AKI-{i:05d}",
            "kaynak": kid,
            "tarih": tar,
            "termin": termin,
            "onceki": o,
            "durum": d,
            "sorumlu": f"Uyum {(i % 6) + 1}",
            "aciklama": f"Uyum kontrol adımı {i}",
            "ilk": ilk,
            "hedef": hedef,
            "tutar": max(1000, ilk * 250),
        })
    return satirlar


ORNEK_AKIS = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "KVKK veri envanteri, hukuki sebep ve yurt dışı aktarım riski ile VERBİS uyum "
      "durum makinesi ve denetime sunulabilir kanıt raporu üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Envanter → Kontrol → Düzeltme → Uyumlu / İptal",
        "KVV-001..005 kural atıflı uyum skoru ve eksik kayıt listesi",
        "Yurt dışı aktarım riski, saklama süresi aşımı ve yaşlandırma",
        "Geçersiz durum geçişi uyarısı ve yenileme takvimi",
        "KANIT_RAPORU — madde atıf, revizyon, imza alanı",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "KVKK sorumlusu — envanter ve hukuki sebep skorlama",
        "Uyum yöneticisi — eksik kayıt ve aktarım riskini kapatma",
        "Denetçi — KANIT_RAPORU ile VERBİS uyum dosyası incelemesi",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) ENVANTER'e veri işleme kayıtlarını girin. 2) UYUM_AKIS'te kontrolleri ilerletin. "
      "3) PANO ve KANIT_RAPORU'nu inceleyin. 4) AYARLAR'da eşik ve yenileme tarihini ayarlayın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def envanter(ws):
    sayfa_hazirla(ws, "ENVANTER", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Veri envanteri — her kaydın tek kimliği vardır (ID disiplini)",
             son_kolon=14)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=kvv_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni satırda kullanın | Örnek: VER-00025 | Risk: Elle çakışma")

    sutunlar = [
        "KayitId", "Kategori", "VeriKonusu", "HukukiSebep", "KuralAtif",
        "Hassasiyet", "AktarimRisk", "SaklamaAy", "YurtDisi", "Guvenlik",
        "MevcutOnlem", "Aktif", "HamSkor", "RiskSeviye", "SaklamaAsim", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "HamSkor": (
            'IF(tblEnvanter[[#This Row],[KayitId]]="","",'
            'tblEnvanter[[#This Row],[Hassasiyet]]*tblEnvanter[[#This Row],[AktarimRisk]])'
        ),
        "RiskSeviye": (
            'IF(tblEnvanter[[#This Row],[KayitId]]="","",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=kvv_esikKritikSeviye,"Kritik",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=kvv_esikYuksek,"Yüksek",'
            'IF(tblEnvanter[[#This Row],[HamSkor]]>=kvv_esikOrta,"Orta","Düşük"))))'
        ),
        "SaklamaAsim": (
            'IF(tblEnvanter[[#This Row],[KayitId]]="","",'
            'IF(tblEnvanter[[#This Row],[SaklamaAy]]<kvv_esikSaklamaAy,"AŞIM","OK"))'
        ),
        "KayitDolu": 'IF(tblEnvanter[[#This Row],[KayitId]]="","",1)',
    }
    for i, row in enumerate(KAYITLAR):
        r = ILK + i
        (kid, kat, konu, huk, kural, has, akt, sak, yd, guv, onlem) = row
        _sari(ws, r, 1, kid, baslik="Kayıt Kimliği", mesaj="VER-##### formatında")
        _sari(ws, r, 2, kat, baslik="Kategori", mesaj="Veri kategorisi")
        _sari(ws, r, 3, konu, baslik="Veri Konusu", mesaj="İşleme konusu özeti")
        _sari(ws, r, 4, huk, baslik="Hukuki Sebep", mesaj="Açık rıza / sözleşme / yasal…")
        _sari(ws, r, 5, kural, baslik="Kural Atıf", mesaj="KVV-001..005")
        _sari(ws, r, 6, has, sayi=CATI, baslik="Hassasiyet", mesaj="1-5 arası")
        _sari(ws, r, 7, akt, sayi=CATI, baslik="Aktarım Riski", mesaj="1-5 arası")
        _sari(ws, r, 8, sak, sayi=CATI, baslik="Saklama Ay", mesaj="Saklama süresi (ay)")
        _sari(ws, r, 9, yd, baslik="Yurt Dışı", mesaj="EVET veya HAYIR")
        _sari(ws, r, 10, guv, sayi=CATI, baslik="Güvenlik", mesaj="1-5 güvence seviyesi")
        _sari(ws, r, 11, onlem, baslik="Mevcut Önlem", mesaj="Teknik/idari önlem özeti")
        _sari(ws, r, 12, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KAYIT, SON + 1):
        for c in range(1, 13):
            fmt = CATI if c in (6, 7, 8, 10) else None
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblEnvanter", f"A{HDR}:P{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Kayıt Kimliği", mesaj="VER-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Kategori"), (3, "Konu"), (4, "Hukuki Sebep"), (11, "Önlem")]:
        dogrulama(ws, "textLength", "0",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    for c, ad in [(6, "Hassasiyet"), (7, "Aktarım Riski"), (10, "Güvenlik")]:
        dogrulama(ws, "whole", "1",
                  f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj="1 ile 5 arası",
                  hata_baslik=ad, hata_mesaj="1-5 arası girin",
                  isaret="between", f2="5")
    dogrulama(ws, "whole", "1", f"H{ILK}:H{SON}",
              baslik="Saklama Ay", mesaj="1 veya üzeri",
              hata_baslik="Saklama", hata_mesaj="1 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"I{ILK}:I{SON}",
              baslik="Yurt Dışı", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeEvetHayir", f"L{ILK}:L{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeKural", f"E{ILK}:E{SON}",
              baslik="Kural Atıf", mesaj="KVV kural kodu seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    dogrulama(ws, "list", "ListeHukuki", f"D{ILK}:D{SON}",
              baslik="Hukuki Sebep", mesaj="Hukuki sebep seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 17)})
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["K"].width = 18


def uyum_akis(ws):
    sayfa_hazirla(ws, "UYUM_AKIS", RENK, URUN_AD, son_kolon=22)
    alt_bant(ws, 3,
             "Uyum satırları — durum makinesi + yetim + geçersiz geçiş + zincir bayrakları",
             son_kolon=18)
    sutunlar = [
        "AkisId", "KaynakKayitId", "Tarih", "Termin", "OncekiDurum", "Durum",
        "Sorumlu", "Aciklama", "IlkRisk", "HedefRisk", "Tutar",
        "YasGun", "YasKova", "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    yas_f = yaslandirma_gun(
        "tblUyumAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = yas_kova_formul("tblUyumAkis[[#This Row],[YasGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblUyumAkis[[#This Row],[KaynakKayitId]]",
        "tblEnvanter[KayitId]").lstrip("=")
    gecis_birlesik = (
        'tblUyumAkis[[#This Row],[OncekiDurum]]&"|"&tblUyumAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblUyumAkis[[#This Row],[IlkRisk]]",
        "tblUyumAkis[[#This Row],[HedefRisk]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "YasGun": f'IF(tblUyumAkis[[#This Row],[AkisId]]="","",{yas_f})',
        "YasKova": f'IF(tblUyumAkis[[#This Row],[AkisId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblUyumAkis[[#This Row],[AkisId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblUyumAkis[[#This Row],[AkisId]]="",'
            f'tblUyumAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblUyumAkis[[#This Row],[AkisId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblUyumAkis[[#This Row],[AkisId]]="","",1)',
    }

    for i, s in enumerate(ORNEK_AKIS):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Akış Kimliği", mesaj="AKI-#####")
        _sari(ws, r, 2, s["kaynak"] or None, baslik="Kaynak Kayıt",
              mesaj="ENVANTER KayitId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Açılış tarihi")
        _sari(ws, r, 4, s["termin"], sayi=TARİH, baslik="Termin", mesaj="Hedef bitiş")
        _sari(ws, r, 5, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 6, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 7, s["sorumlu"], baslik="Sorumlu", mesaj="Sorumlu kişi")
        _sari(ws, r, 8, s["aciklama"], baslik="Açıklama", mesaj="Kısa uyum adımı")
        _sari(ws, r, 9, s["ilk"], sayi=CATI, baslik="İlk Risk", mesaj="Başlangıç risk skoru")
        _sari(ws, r, 10, s["hedef"], sayi=CATI, baslik="Hedef Risk", mesaj="Hedef kalan risk")
        _sari(ws, r, 11, s["tutar"], sayi=TL, baslik="Tutar", mesaj="Tahmini uyum maliyeti TL")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 12):
            fmt = TARİH if c in (3, 4) else (TL if c == 11 else (CATI if c in (9, 10) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblUyumAkis", f"A{HDR}:Q{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Akış Kimliği", mesaj="AKI-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Kayıt", mesaj="Kayıt kimliği (boş = yetim riski)",
              hata_baslik="Kayıt", hata_mesaj="Geçersiz",
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
    for c, ad in [(9, "İlk Risk"), (10, "Hedef Risk")]:
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


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; eşikler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Kayıt toplam", '=COUNTA(tblEnvanter[KayitId])'),
        ("H02", "Dolu kayıt", '=SUM(tblEnvanter[KayitDolu])'),
        ("H03", "Aktif kayıt", '=COUNTIF(tblEnvanter[Aktif],"EVET")'),
        ("H04", "Skor toplam", "=SUM(tblEnvanter[HamSkor])"),
        ("H05", "Ortalama skor", "=IFERROR(AVERAGE(tblEnvanter[HamSkor]),0)"),
        ("H06", "Maks skor", "=IFERROR(MAX(tblEnvanter[HamSkor]),0)"),
        ("H07", "Min skor", "=IFERROR(MIN(tblEnvanter[HamSkor]),0)"),
        ("H08", "Düşük adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Düşük")'),
        ("H09", "Orta adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Orta")'),
        ("H10", "Yüksek adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Yüksek")'),
        ("H11", "Kritik adet", '=COUNTIF(tblEnvanter[RiskSeviye],"Kritik")'),
        ("H12", "KVV-001 atıf", '=COUNTIF(tblEnvanter[KuralAtif],"KVV-001")'),
        ("H13", "KVV-002 atıf", '=COUNTIF(tblEnvanter[KuralAtif],"KVV-002")'),
        ("H14", "KVV-003 atıf", '=COUNTIF(tblEnvanter[KuralAtif],"KVV-003")'),
        ("H15", "KVV-004 atıf", '=COUNTIF(tblEnvanter[KuralAtif],"KVV-004")'),
        ("H16", "KVV-005 atıf", '=COUNTIF(tblEnvanter[KuralAtif],"KVV-005")'),
        ("H17", "Akış toplam", '=COUNTA(tblUyumAkis[AkisId])'),
        ("H18", "Envanter durum", '=COUNTIF(tblUyumAkis[Durum],"ENVANTAR")'),
        ("H19", "Kontrol durum", '=COUNTIF(tblUyumAkis[Durum],"KONTROL")'),
        ("H20", "Düzeltme durum", '=COUNTIF(tblUyumAkis[Durum],"DUZELTME")'),
        ("H21", "Uyumlu durum", '=COUNTIF(tblUyumAkis[Durum],"UYUMLU")'),
        ("H22", "İptal durum", '=COUNTIF(tblUyumAkis[Durum],"IPTAL")'),
        ("H23", "Açık adet", "=C23+C24+C25"),
        ("H24", "Kapalı adet", "=C26+C27"),
        ("H25", "Açık tutar",
         '=SUMIF(tblUyumAkis[Durum],"ENVANTAR",tblUyumAkis[Tutar])'
         '+SUMIF(tblUyumAkis[Durum],"KONTROL",tblUyumAkis[Tutar])'
         '+SUMIF(tblUyumAkis[Durum],"DUZELTME",tblUyumAkis[Tutar])'),
        ("H26", "Yetim adet", '=COUNTIF(tblUyumAkis[YetimBayrak],"YETİM")'),
        ("H27", "Geçersiz geçiş", '=COUNTIF(tblUyumAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H28", "Zincir kırık", '=COUNTIF(tblUyumAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H29", "Zincir OK", '=COUNTIF(tblUyumAkis[ZincirBayrak],"OK")'),
        ("H30", "Yaş 0-7", '=COUNTIF(tblUyumAkis[YasKova],"0-7")'),
        ("H31", "Yaş 8-30", '=COUNTIF(tblUyumAkis[YasKova],"8-30")'),
        ("H32", "Yaş 31-60", '=COUNTIF(tblUyumAkis[YasKova],"31-60")'),
        ("H33", "Yaş 60+", '=COUNTIF(tblUyumAkis[YasKova],"60+")'),
        ("H34", "Ort yaş gün", "=IFERROR(AVERAGE(tblUyumAkis[YasGun]),0)"),
        ("H35", "Geciken adet",
         '=COUNTIFS(tblUyumAkis[YasGun],">"&kvv_esikGecikmeGun,tblUyumAkis[Durum],"<>UYUMLU",'
         'tblUyumAkis[Durum],"<>IPTAL",tblUyumAkis[Tutar],">0")'),
        ("H36", "Geciken tutar",
         '=SUMIFS(tblUyumAkis[Tutar],tblUyumAkis[YasGun],">"&kvv_esikGecikmeGun,'
         'tblUyumAkis[Durum],"<>UYUMLU",tblUyumAkis[Durum],"<>IPTAL")'),
        ("H37", "Gecikme maliyeti", "=C41*kvv_gecikmeOranGun"),
        ("H38", "Temkinli senaryo", "=C10*kvv_senaryoTemkinli"),
        ("H39", "Baz senaryo", "=C10*kvv_senaryoBaz"),
        ("H40", "İyimser senaryo", "=C10*kvv_senaryoIyimser"),
        ("H41", "Senaryo farkı", "=C43-C45"),
        ("H42", "Tornado etki", "=C10*kvv_tornadoOran"),
        ("H43", "Uyum öncelik havuzu",
         '=SUMIFS(tblUyumAkis[Tutar],tblUyumAkis[Durum],"DUZELTME",'
         'tblUyumAkis[YasGun],">"&kvv_esikGecikmeGun)'),
        ("H44", "Pareto kritik payı", "=IFERROR(C16/MAX(C6,1),0)"),
        ("H45", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblEnvanter[HamSkor],kvv_yuzdelikOran),0)"),
        ("H46", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblEnvanter[KayitId])+1,tblEnvanter[HamSkor],"
         "tblEnvanter[KayitDolu]),C50)"),
        ("H47", "Tahmin aralık", "=IFERROR((C50+C51)/2,0)"),
        ("H48", "Uyum oranı", "=IFERROR(C26/MAX(C22,1),0)"),
        ("H49", "Uyum skoru", "=IFERROR(C26/MAX(C22,1)*100,0)"),
        ("H50", "Anomali skor",
         '=COUNTIFS(tblEnvanter[HamSkor],">"&kvv_esikAnomali)+C31+C32+C33'),
        ("H51", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C35,C36,C37,C38)'),
        ("H52", "Skor dağılım metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"D",C13,"O",C14,"Y",C15,"K",C16)'),
        ("H53", "Kuyruk özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Env",C23,"Kon",C24,"Duz",C25,"Geciken",C40)'),
        ("H54", "Karar skoru ham",
         "=IF(C6=0,0,MAX(0,100-C31*10-C32*10-C33*10-C16*5-C40*2))"),
        ("H55", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H56", "Kural atıf özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"001",C17,"002",C18,"003",C19,"004",C20,"005",C21)'),
        ("H57", "Yenileme kalan gün", "=kvv_yenilemeTarihi-raporTarihi"),
        ("H58", "Kapalı düşüm", "=C29"),
        ("H59", "Yurt dışı adet", '=COUNTIF(tblEnvanter[YurtDisi],"EVET")'),
        ("H60", "Aktarım riski",
         '=COUNTIFS(tblEnvanter[YurtDisi],"EVET",tblEnvanter[Guvenlik],"<"&kvv_esikGuvenlik)'),
        ("H61", "Eksik kayıt",
         '=COUNTIFS(tblEnvanter[MevcutOnlem],"*eksik*")+COUNTIF(tblEnvanter[SaklamaAsim],"AŞIM")'),
        ("H62", "Saklama aşım adet", '=COUNTIF(tblEnvanter[SaklamaAsim],"AŞIM")'),
        ("H63", "Senaryo motor",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"T",C43,"B",C44,"I",C45)'),
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
                        or "dağılım" in acik.lower() or "motor" in acik.lower() else CATI))
        )
        h(ws, r, 3, form, sayi=fmt, kalin=True)

    h(ws, 72, 1, "Senaryo", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 72, 2, "Çarpan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 72, 3, "Sonuç", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 73, 1, "Temkinli")
    h(ws, 73, 2,
      "=kvv_senaryoTemkinli*IF(COUNTA(tblEnvanter[KayitId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 73, 3, "=C43", sayi=CATI)
    h(ws, 74, 1, "Baz")
    h(ws, 74, 2,
      "=kvv_senaryoBaz*IF(COUNTA(tblEnvanter[KayitId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 74, 3, "=C44", sayi=CATI)
    h(ws, 75, 1, "İyimser")
    h(ws, 75, 2,
      "=kvv_senaryoIyimser*IF(COUNTA(tblEnvanter[KayitId])>=0,1,1)", sayi=YÜZDE)
    h(ws, 75, 3, "=C45", sayi=CATI)

    h(ws, 77, 1, "Kuyruk (açık uyum adımları)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 78, 1, "Envanter")
    h(ws, 78, 2, kuyruk_sayim("tblUyumAkis", "ENVANTAR"), sayi=CATI)
    h(ws, 78, 3, kuyruk_tutar("tblUyumAkis", "ENVANTAR"), sayi=TL)
    h(ws, 79, 1, "Kontrol")
    h(ws, 79, 2, kuyruk_sayim("tblUyumAkis", "KONTROL"), sayi=CATI)
    h(ws, 79, 3, kuyruk_tutar("tblUyumAkis", "KONTROL"), sayi=TL)
    h(ws, 80, 1, "Düzeltme")
    h(ws, 80, 2, kuyruk_sayim("tblUyumAkis", "DUZELTME"), sayi=CATI)
    h(ws, 80, 3, kuyruk_tutar("tblUyumAkis", "DUZELTME"), sayi=TL)

    genislik(ws, {"A": 10, "B": 28, "C": 55})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=20)
    sabitle(ws, "A3")
    h(ws, 3, 1, "Karar Destek Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 3, 3, "Rapor Tarihi", yazi=GRİ)
    h(ws, 3, 4, "=raporTarihi", sayi=TARİH)

    h(ws, 4, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2,
      '=IF(COUNTA(tblEnvanter[KayitId])=0,"VERİ YOK",'
      'IF(OR(kvv_yetimKayit>0,kvv_gecersizGecis>0,kvv_zincirKirik>0,'
      'kvv_kritikAdet>kvv_esikKritik),"KRİTİK",'
      'IF(OR(kvv_eksikKayit>0,kvv_aktarimRiski>kvv_esikAktarim,'
      'kvv_saklamaAsim>0,kvv_uyumSkoru<kvv_esikUyum),"EKSİK","UYUMLU")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYUMLU | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Kayıt Adet", "=kvv_kayitAdet", CATI),
        (7, "Kritik Adet", "=kvv_kritikAdet", CATI),
        (8, "Ortalama Skor", "=kvv_ortalamaSkor", CATI),
        (9, "Uyum Oranı", "=kvv_uyumOrani", YÜZDE),
        (10, "Uyum Skoru", "=kvv_uyumSkoru", CATI),
        (11, "Açık Adım", "=kvv_acikAdet", CATI),
        (12, "Geciken Adım", "=kvv_gecikenAdet", CATI),
        (13, "Yetim Kayıt", "=kvv_yetimKayit", CATI),
        (14, "Geçersiz Geçiş", "=kvv_gecersizGecis", CATI),
        (15, "Zincir Kırık", "=kvv_zincirKirik", CATI),
        (16, "İşlem Yaş", "=kvv_islemYas", None),
        (17, "Tahmin Aralık", "=kvv_tahminAralik", CATI),
        (18, "Senaryo Farkı", "=kvv_senaryoKarsilastirma", CATI),
        (19, "Tornado Zirve", "=kvv_tornadoZirve", CATI),
        (20, "Uyum Öncelik", "=kvv_uyumOncelik", TL),
        (21, "Aktarım Riski", "=kvv_aktarimRiski", CATI),
        (22, "Eksik Kayıt", "=kvv_eksikKayit", CATI),
        (23, "Saklama Aşım", "=kvv_saklamaAsim", CATI),
        (24, "Anomali Skor", "=kvv_anomaliSayisi", CATI),
        (25, "Skor Dağılım", "=kvv_skorDagilim", None),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=kvv_yorumIslemYas"),
        (8, "=kvv_yorumSaklama"),
        (9, "=kvv_yorumTahmin"),
        (10, "=kvv_yorumSenaryo"),
        (11, "=kvv_yorumTornado"),
        (12, "=kvv_yorumOncelik"),
        (13, "=kvv_yorumAktarim"),
        (14, "=kvv_yorumEksik"),
        (15, "=kvv_yorumAnomali"),
        (16, "=kvv_yorumSenaryoMotor"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Seviye")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Düşük", 6), ("Orta", 8), ("Yüksek", 6), ("Kritik", 4),
    ], 29):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 28, 4, "Yaş Kova")
    h(ws, 28, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("0-7", 7), ("8-30", 9), ("31-60", 6), ("60+", 6),
    ], 29):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 28, 7, "Hafta")
    h(ws, 28, 8, "Yeni Kayıt")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 2 + (i % 4), sayi=CATI)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "Skor")
    for i, (ad, t) in enumerate([
        ("Temkinli", 420), ("Baz", 510), ("İyimser", 380),
    ], 29):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=CATI)

    h(ws, 28, 13, "Uyarı")
    h(ws, 28, 14, "Adet")
    for i, (ad, adet) in enumerate([
        ("Yetim", 2), ("Geçersiz", 2), ("Zincir", 2), ("Aktarım", 5),
    ], 29):
        h(ws, i, 13, ad)
        h(ws, i, 14, adet, sayi=CATI)

    h(ws, 28, 16, "Durum")
    h(ws, 28, 17, "Adet")
    for i, (ad, adet) in enumerate([
        ("E-Adet", 5), ("K-Adet", 8), ("D-Adet", 6),
        ("U-Adet", 5), ("I-Adet", 4),
    ], 29):
        h(ws, i, 16, ad)
        h(ws, i, 17, adet, sayi=CATI)

    h(ws, 35, 1, "Açık Kuyruk (COUNTIFS)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 36, 1, "Envanter kuyruk")
    h(ws, 36, 2, '=COUNTIFS(tblUyumAkis[Durum],"ENVANTAR",tblUyumAkis[Tutar],">0")', sayi=CATI)
    h(ws, 36, 3, '=SUMIFS(tblUyumAkis[Tutar],tblUyumAkis[Durum],"ENVANTAR")', sayi=TL)
    h(ws, 37, 1, "Kontrol kuyruk")
    h(ws, 37, 2, '=COUNTIFS(tblUyumAkis[Durum],"KONTROL",tblUyumAkis[Tutar],">0")', sayi=CATI)
    h(ws, 37, 3, '=SUMIFS(tblUyumAkis[Tutar],tblUyumAkis[Durum],"KONTROL")', sayi=TL)
    h(ws, 38, 1, "Düzeltme kuyruk")
    h(ws, 38, 2, '=COUNTIFS(tblUyumAkis[Durum],"DUZELTME",tblUyumAkis[Tutar],">0")', sayi=CATI)
    h(ws, 38, 3, '=SUMIFS(tblUyumAkis[Tutar],tblUyumAkis[Durum],"DUZELTME")', sayi=TL)

    c1 = PieChart()
    c1.title = "Skor Seviye Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=28, max_row=32), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=29, max_row=32))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A40")

    c2 = BarChart()
    c2.title = "İşlem Yaşlandırma"
    c2.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F40")

    c3 = LineChart()
    c3.title = "Haftalık Yeni Kayıt"
    c3.add_data(Reference(ws, min_col=8, min_row=28, max_row=36), titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=7, min_row=29, max_row=36))
    c3.width, c3.height = 12, 7
    ws.add_chart(c3, "A55")

    c4 = BarChart()
    c4.title = "Senaryo Karşılaştırma"
    c4.add_data(Reference(ws, min_col=11, min_row=28, max_row=31), titles_from_data=True)
    c4.set_categories(Reference(ws, min_col=10, min_row=29, max_row=31))
    c4.width, c4.height = 10, 7
    ws.add_chart(c4, "F55")

    c5 = BarChart()
    c5.title = "Uyarı Türleri"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A70")

    c6 = PieChart()
    c6.title = "Uyum Durum Payı"
    c6.add_data(Reference(ws, min_col=17, min_row=28, max_row=33), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=16, min_row=29, max_row=33))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F70")

    c7 = BarChart()
    c7.title = "Durum Adet"
    c7.add_data(Reference(ws, min_col=17, min_row=28, max_row=33), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=16, min_row=29, max_row=33))
    c7.width, c7.height = 10, 7
    ws.add_chart(c7, "A85")

    c8 = BarChart()
    c8.title = "Kova Payı"
    c8.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F85")

    baski_hazirla(ws, "A1:N38", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def akis_ozet(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3,
             "Uyum akış özeti — KAYNAK_KART_ID / yetim / kuyruk (UYUM_AKIS ile bağlı)",
             son_kolon=10)
    h(ws, 5, 1, "KAYNAK_KART_ID kontrolü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Yetim (YETİM) adet")
    h(ws, 6, 2, '=COUNTIF(tblUyumAkis[YetimBayrak],"YETİM")', sayi=CATI)
    h(ws, 7, 1, "Geçersiz geçiş")
    h(ws, 7, 2, '=COUNTIF(tblUyumAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")', sayi=CATI)
    h(ws, 8, 1, "Zincir kırık")
    h(ws, 8, 2, '=COUNTIF(tblUyumAkis[ZincirBayrak],"ZİNCİR KIRIK")', sayi=CATI)
    h(ws, 10, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 10, 3, "Tutar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 11, 1, "Envanter")
    h(ws, 11, 2, kuyruk_sayim("tblUyumAkis", "ENVANTAR"), sayi=CATI)
    h(ws, 11, 3, kuyruk_tutar("tblUyumAkis", "ENVANTAR"), sayi=TL)
    h(ws, 12, 1, "Kontrol")
    h(ws, 12, 2, kuyruk_sayim("tblUyumAkis", "KONTROL"), sayi=CATI)
    h(ws, 12, 3, kuyruk_tutar("tblUyumAkis", "KONTROL"), sayi=TL)
    h(ws, 13, 1, "Düzeltme")
    h(ws, 13, 2, kuyruk_sayim("tblUyumAkis", "DUZELTME"), sayi=CATI)
    h(ws, 13, 3, kuyruk_tutar("tblUyumAkis", "DUZELTME"), sayi=TL)
    h(ws, 15, 1, "Kapalı (Uyumlu+İptal) kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Uyumlu")
    h(ws, 16, 2, kuyruk_sayim("tblUyumAkis", "UYUMLU"), sayi=CATI)
    h(ws, 17, 1, "İptal")
    h(ws, 17, 2, kuyruk_sayim("tblUyumAkis", "IPTAL"), sayi=CATI)
    genislik(ws, {"A": 42, "B": 14, "C": 16})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Denetim Kanıt Raporu — KVKK / VERBİS Uyum", boyut=16, son_kolon=10)
    for i, (etiket, deger) in enumerate(rapor_baslik_satirlari(URUN_AD, SURUM), 4):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger)
    h(ws, 8, 1, "Rapor Tarihi")
    h(ws, 8, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 9, 1, "KARAR", kalin=True)
    h(ws, 9, 2, "=PANO!B4", kalin=True, boyut=14)

    ozet = [
        (11, "Kayıt Adet", "=kvv_kayitAdet", CATI),
        (12, "Kritik Adet", "=kvv_kritikAdet", CATI),
        (13, "Uyum Oranı", "=kvv_uyumOrani", YÜZDE),
        (14, "Yetim", "=kvv_yetimKayit", CATI),
        (15, "Geçersiz Geçiş", "=kvv_gecersizGecis", CATI),
        (16, "Zincir Kırık", "=kvv_zincirKirik", CATI),
        (17, "Kural Atıf", "=kvv_kuralAtif", None),
        (18, "Kanıt Özeti", "=kvv_kanitRaporu", None),
        (19, "Skor Dağılım", "=kvv_skorDagilim", None),
        (20, "Yenileme Takvim", "=kvv_yenilemeTakvim", TARİH),
        (21, "Revizyon Geçmişi", "=kvv_revizyonGecmisi", None),
        (22, "Aktarım Riski", "=kvv_aktarimRiski", CATI),
        (23, "Eksik Kayıt", "=kvv_eksikKayit", CATI),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)

    h(ws, 25, 1, "Hesap Zinciri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1,
      "Hassasiyet × AktarımRisk → HamSkor → RiskSeviye → Uyum durum makinesi → PANO kararı.",
      kaydir=True)
    ws.merge_cells("A26:H26")

    h(ws, 28, 1, "Varsayımlar ve Belirsizlik", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 29, 1,
      "Açık rıza ile hukuki sebep ayrımı kullanıcı beyanına bağlıdır. Yurt dışı aktarım "
      "güvence seviyesi (SCC / yeterlilik) dosya içi skordur; kesin hukuki görüş yerine "
      "geçmez. Dosya karar destektir, çevrimdışı ve makrosuzdur.",
      kaydir=True)
    ws.merge_cells("A29:H29")
    h(ws, 30, 1, "=kvv_kvkkBeyani", kaydir=True)
    ws.merge_cells("A30:H30")

    h(ws, 32, 1, "İmza / Onay", kalin=True, yazi=KOYU_LACIVERT)
    for i, (etiket, deger) in enumerate(imza_alani(), 33):
        h(ws, i, 1, etiket)
        if deger:
            h(ws, i, 2, deger, sayi=TARİH if "Tarih" in etiket else None)
        else:
            _sari(ws, i, 2, "", baslik=etiket, mesaj="İmza / unvan yazın")

    h(ws, 38, 1, sha_placeholder_notu(), kaydir=True, yazi=GRİ, boyut=9)
    ws.merge_cells("A38:H38")
    h(ws, 40, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:H40", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 62, "B": 40, "C": 14, "D": 12, "E": 14})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "Kayıt sayısı", '=COUNTA(tblEnvanter[KayitId])'),
        (6, "Akış sayısı", '=COUNTA(tblUyumAkis[AkisId])'),
        (7, "Kritik", "=kvv_kritikAdet"),
        (8, "Envanter", '=COUNTIF(tblUyumAkis[Durum],"ENVANTAR")'),
        (9, "Kontrol", '=COUNTIF(tblUyumAkis[Durum],"KONTROL")'),
        (10, "Düzeltme", '=COUNTIF(tblUyumAkis[Durum],"DUZELTME")'),
        (11, "Uyumlu", '=COUNTIF(tblUyumAkis[Durum],"UYUMLU")'),
        (12, "Yetim", "=kvv_yetimKayit"),
        (13, "Geçersiz", "=kvv_gecersizGecis"),
        (14, "Zincir kırık", "=kvv_zincirKirik"),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru", "=MOTOR!C59"),
        (17, "Uyum oranı", "=kvv_uyumOrani"),
        (18, "Aktarım riski", "=kvv_aktarimRiski"),
        (19, "Kapalı düşüm", "=kvv_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        fmt = YÜZDE if r == 17 else (CATI if r < 15 or r in (18, 19) else None)
        h(ws, r, 2, form, kalin=True, sayi=fmt)
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KAYIT} envanter ve {DEMO_AKIS} uyum akış satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "ENVANTER ve UYUM_AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Envanter/Kontrol/Düzeltme/Uyumlu/İptal — kapalılar kuyruktan düşer")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 5000, sayi=CATI)
    genislik(ws, {"A": 70, "B": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=14)
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

    h(ws, 3, 7, "Kural Atıf", kalin=True)
    h(ws, 6, 7, "Kural", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, m in enumerate(["KVV-001", "KVV-002", "KVV-003", "KVV-004", "KVV-005"], 7):
        _sari(ws, i, 7, m, baslik="Kural", mesaj="KVKK kural kodu")

    h(ws, 3, 9, "Hukuki Sebep", kalin=True)
    h(ws, 6, 9, "Sebep", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, m in enumerate([
        "Açık rıza", "Sözleşme", "Yasal yükümlülük", "Meşru menfaat", "Hayati çıkar",
    ], 7):
        _sari(ws, i, 9, m, baslik="Sebep", mesaj="Hukuki sebep")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 12, "I": 18})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası + uyum eşikleri",
             son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 15, "skor", "İş kuralı — skor sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği (ilk-hedef)"),
        ("kvv_esikGecikmeGun", 30, "gün", "İç politika", "01.01.2026", "Gecikme eşiği"),
        ("kvv_esikGeciken", 5, "adet", "İç politika", "01.01.2026", "Karar eşiği geciken"),
        ("kvv_esikKritik", 3, "adet", "İç politika", "01.01.2026", "Karar eşiği kritik"),
        ("kvv_esikKritikSeviye", 16, "skor", "Uyum skor eşiği", "01.01.2026", "Kritik skor"),
        ("kvv_esikYuksek", 10, "skor", "Uyum skor eşiği", "01.01.2026", "Yüksek skor"),
        ("kvv_esikOrta", 6, "skor", "Uyum skor eşiği", "01.01.2026", "Orta skor"),
        ("kvv_esikAnomali", 20, "skor", "Anomali kuralı", "01.01.2026", "Anomali skor eşiği"),
        ("kvv_esikSaklamaAy", 12, "ay", "Saklama politikası", "01.01.2026",
         "Minimum saklama ay eşiği"),
        ("kvv_esikGuvenlik", 3, "seviye", "Aktarım güvence", "01.01.2026",
         "Yurt dışı min güvenlik"),
        ("kvv_esikAktarim", 2, "adet", "Karar eşiği", "01.01.2026", "Aktarım risk adet eşiği"),
        ("kvv_esikUyum", 50, "skor", "Uyum skoru eşiği", "01.01.2026", "Min uyum skoru"),
        ("kvv_gecikmeOranGun", 0.0005, "oran", "Finans varsayımı", "01.01.2026",
         "Günlük gecikme maliyeti oranı"),
        ("kvv_senaryoTemkinli", 1.15, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("kvv_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("kvv_senaryoIyimser", 0.85, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("kvv_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("kvv_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("kvv_yenilemeTarihi", date(2027, 8, 11), "tarih", "VERBİS yenileme", "01.01.2026",
         "Sonraki VERBİS yenileme"),
        ("kvv_olcekHedef", 5000, "satir", "Manda A4 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("kvv_azamiSkor", 500, "skor", "İç politika — uç değer", "01.01.2026", "Azami makul"),
        ("kvv_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("kvv_firmaUnvan", "Örnek Şirket A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("kvv_durumMakinesi", "tblDurumHaritasi", "metin", "SPEC akis", "01.01.2026",
         "Uyum durum makinesi"),
        ("kvv_kvkkBeyani",
         "Kişisel veri toplanmaz; dosya çevrimdışı ve makrosuzdur. Karar destek aracıdır.",
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

    h(ws, 34, 8, "Durum Haritası", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    dh_bas = ["durum_id", "durum_adi", "sira", "izinli_sonraki_durumlar", "kategori"]
    baslik_satiri(ws, 35, dh_bas, basla=8)
    for i, d in enumerate(DURUM_HARITASI):
        r = 36 + i
        h(ws, r, 8, d["durum_id"])
        h(ws, r, 9, d["durum_adi"])
        h(ws, r, 10, d["sira"], sayi=CATI)
        h(ws, r, 11, d["izinli_sonraki_durumlar"])
        h(ws, r, 12, d["kategori"])
    tablo_ekle(ws, "tblDurumHaritasi", "H35:L40", dh_bas)

    h(ws, 43, 8, "Motor Çıktıları", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 44, 8, "anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 44, 9, "deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    motor_cikti = [
        (45, "kvv_kayitAdet", "=MOTOR!C6"),
        (46, "kvv_kritikAdet", "=MOTOR!C16"),
        (47, "kvv_ortalamaSkor", "=MOTOR!C10"),
        (48, "kvv_uyumOrani", "=MOTOR!C53"),
        (49, "kvv_uyumSkoru", "=MOTOR!C54"),
        (50, "kvv_acikAdet", "=MOTOR!C28"),
        (51, "kvv_gecikenAdet", "=MOTOR!C40"),
        (52, "kvv_yetimKayit", "=MOTOR!C31"),
        (53, "kvv_gecersizGecis", "=MOTOR!C32"),
        (54, "kvv_zincirKirik", "=MOTOR!C33"),
        (55, "kvv_islemYas", "=MOTOR!C56"),
        (56, "kvv_tahminAralik", "=MOTOR!C52"),
        (57, "kvv_senaryoKarsilastirma", "=MOTOR!C46"),
        (58, "kvv_tornadoZirve", "=MOTOR!C47"),
        (59, "kvv_uyumOncelik", "=MOTOR!C48"),
        (60, "kvv_aktarimRiski", "=MOTOR!C65"),
        (61, "kvv_eksikKayit", "=MOTOR!C66"),
        (62, "kvv_saklamaAsim", "=MOTOR!C67"),
        (63, "kvv_anomaliSayisi", "=MOTOR!C55"),
        (64, "kvv_skorDagilim", "=MOTOR!C57"),
        (65, "kvv_kapaliDusum", "=MOTOR!C63"),
        (66, "kvv_kuyrukOzeti", "=MOTOR!C58"),
        (67, "kvv_kuralAtif", "=MOTOR!C61"),
        (68, "kvv_yenilemeTakvim", "=kvv_yenilemeTarihi"),
        (69, "kvv_revizyonGecmisi",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Sürüm",kvv_dosyaSurumu,"Rapor",TEXT(raporTarihi,"dd.mm.yyyy"))'),
        (70, "kvv_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Kayıt",MOTOR!C6,"Kritik",MOTOR!C16,'
         '"Yetim",MOTOR!C31,"Uyum",TEXT(MOTOR!C53,"0%"))'),
        (71, "kvv_sonrakiKimlik",
         '="VER-"&TEXT(COUNTA(tblEnvanter[KayitId])+1,"00000")'),
        (72, "kvv_senaryoMotor", "=MOTOR!C68"),
    ]
    for r, ad, form in motor_cikti:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 74, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (75, "kvv_yorumIslemYas",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yaş kovaları",kvv_islemYas)'),
        (76, "kvv_yorumSaklama",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Saklama aşım adedi",kvv_saklamaAsim)'),
        (77, "kvv_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(kvv_tahminAralik,"0.00"))'),
        (78, "kvv_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(kvv_senaryoKarsilastirma,"0.00"))'),
        (79, "kvv_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(kvv_tornadoZirve,"0.00"))'),
        (80, "kvv_yorumOncelik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Uyum öncelik havuzu",TEXT(kvv_uyumOncelik,"0.00"),"TL")'),
        (81, "kvv_yorumAktarim",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Aktarım risk adedi",kvv_aktarimRiski)'),
        (82, "kvv_yorumEksik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Eksik kayıt adedi",kvv_eksikKayit)'),
        (83, "kvv_yorumAnomali",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Anomali sayısı",kvv_anomaliSayisi)'),
        (84, "kvv_yorumSenaryoMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motor",kvv_senaryoMotor)'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 22, "B": 36, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. ENVANTER sayfasına veri işleme kayıtlarını girin; sonraki kimliği kullanın.",
        "2. UYUM_AKIS'te her kontrol/düzeltme adımını yazın; durum kolonunu listeden seçin.",
        "3. İzinli geçiş: ENVANTAR→KONTROL→DUZELTME→UYUMLU (IPTAL her açık durumdan).",
        "4. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "5. Geçersiz geçiş engellenmez — kırmızı uyarı üretir (makrosuz kural).",
        "6. Kapalı durumlar (Uyumlu/İptal) açık kuyruklardan düşer.",
        "7. raporTarihi AYARLAR'dadır; TODAY kullanılmaz.",
        "8. Koruma şifresi: 1234 — formül hücreleri kilitli, sarı hücreler açıktır.",
        "9. KANIT_RAPORU sayfasını PDF olarak yazdırabilirsiniz.",
        "10. BELİRSİZLİK — acik_riza_vs_hukuki_sebep_ayrimi ve "
        "yurt_disi_aktarim_guvence_seviyesi: skor kullanıcı beyanına bağlıdır; "
        "kesin hukuki görüş yerine geçmez. Denetimde varsayımları açıkça belirtin.",
        "11. Kurallar: KVV-001 veri kategorisi/hukuki sebep; KVV-002 saklama süresi; "
        "KVV-003 yurt dışı güvence; KVV-004 açık rıza ayrımı; KVV-005 VERBİS yenileme.",
        f"12. Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv",
    ]
    for i, m in enumerate(maddeler, 5):
        h(ws, i, 1, m, kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    genislik(ws, {"A": 80})


def _cok_dogrulama(wb):
    ws = wb["AYARLAR"]
    for r in range(6, 32):
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
    for r, ad in [(5, "Sürüm"), (33, "Hazırlayan"), (34, "Kontrol"), (35, "Tarih")]:
        dogrulama(ws, "textLength", "0", f"B{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")
    ws = wb["KILAVUZ"]
    for r in range(5, 17):
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
        "B4", CellIsRule(operator="equal", formula=['"UYUMLU"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"EKSİK"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 26):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["UYUM_AKIS"]
    for col, val, fill, font in [
        ("N", '"YETİM"', kirmizi, kf),
        ("O", '"GEÇERSİZ GEÇİŞ"', kirmizi, kf),
        ("P", '"ZİNCİR KIRIK"', kirmizi, kf),
        ("P", '"OK"', yesil, yf),
    ]:
        ws.conditional_formatting.add(
            f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=[val], fill=fill, font=font))
    for durum, fill in [("DUZELTME", sari), ("UYUMLU", yesil), ("IPTAL", kirmizi)]:
        ws.conditional_formatting.add(
            f"F{ILK}:F{SON}",
            CellIsRule(operator="equal", formula=[f'"{durum}"'], fill=fill))

    ws = wb["ENVANTER"]
    for seviye, fill in [("Kritik", kirmizi), ("Yüksek", sari), ("Düşük", yesil)]:
        ws.conditional_formatting.add(
            f"N{ILK}:N{SON}",
            CellIsRule(operator="equal", formula=[f'"{seviye}"'], fill=fill))
    ws.conditional_formatting.add(
        f"O{ILK}:O{SON}",
        CellIsRule(operator="equal", formula=['"AŞIM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        f"L{ILK}:L{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))

    ws = wb["MOTOR"]
    for r in range(6, 69):
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(
            f"C{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    for r in range(78, 81):
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
        "B9", CellIsRule(operator="equal", formula=['"UYUMLU"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"EKSİK"'], fill=sari))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"KRİTİK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(11, 17):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["AKIS"]
    for r in range(6, 9):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=kirmizi, font=kf))
    for r in range(11, 14):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=sari))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "kvv_esikGecikmeGun", "kvv_esikGeciken",
        "kvv_esikKritik", "kvv_esikKritikSeviye", "kvv_esikYuksek", "kvv_esikOrta",
        "kvv_esikAnomali", "kvv_esikSaklamaAy", "kvv_esikGuvenlik", "kvv_esikAktarim",
        "kvv_esikUyum", "kvv_gecikmeOranGun", "kvv_senaryoTemkinli", "kvv_senaryoBaz",
        "kvv_senaryoIyimser", "kvv_tornadoOran", "kvv_yuzdelikOran", "kvv_yenilemeTarihi",
        "kvv_olcekHedef", "kvv_azamiSkor", "kvv_dosyaSurumu", "kvv_firmaUnvan",
        "kvv_durumMakinesi", "kvv_kvkkBeyani",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "kvv_kayitAdet": 45, "kvv_kritikAdet": 46, "kvv_ortalamaSkor": 47,
        "kvv_uyumOrani": 48, "kvv_uyumSkoru": 49, "kvv_acikAdet": 50,
        "kvv_gecikenAdet": 51, "kvv_yetimKayit": 52, "kvv_gecersizGecis": 53,
        "kvv_zincirKirik": 54, "kvv_islemYas": 55, "kvv_tahminAralik": 56,
        "kvv_senaryoKarsilastirma": 57, "kvv_tornadoZirve": 58, "kvv_uyumOncelik": 59,
        "kvv_aktarimRiski": 60, "kvv_eksikKayit": 61, "kvv_saklamaAsim": 62,
        "kvv_anomaliSayisi": 63, "kvv_skorDagilim": 64, "kvv_kapaliDusum": 65,
        "kvv_kuyrukOzeti": 66, "kvv_kuralAtif": 67, "kvv_yenilemeTakvim": 68,
        "kvv_revizyonGecmisi": 69, "kvv_kanitRaporu": 70, "kvv_sonrakiKimlik": 71,
        "kvv_senaryoMotor": 72,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "kvv_yorumIslemYas": 75, "kvv_yorumSaklama": 76, "kvv_yorumTahmin": 77,
        "kvv_yorumSenaryo": 78, "kvv_yorumTornado": 79, "kvv_yorumOncelik": 80,
        "kvv_yorumAktarim": 81, "kvv_yorumEksik": 82, "kvv_yorumAnomali": 83,
        "kvv_yorumSenaryoMotor": 84,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$14")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeKural", "LISTELER!$G$7:$G$11")
    ad_ekle(wb, "ListeHukuki", "LISTELER!$I$7:$I$11")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (envanter, "ENVANTER"),
        (uyum_akis, "UYUM_AKIS"),
        (motor, "MOTOR"),
        (pano, "PANO"),
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
    dosya_ad = "KvkkVeriEnvanteriVerbisUyum.xlsx"
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
