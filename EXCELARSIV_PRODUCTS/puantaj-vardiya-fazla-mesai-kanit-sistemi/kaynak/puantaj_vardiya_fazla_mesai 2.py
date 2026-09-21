#!/usr/bin/env python3
"""Puantaj-Vardiya-Fazla Mesai Kanıt Sistemi — A3+A4 üretim betiği (manda v6)."""

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

URUN_AD = "Puantaj-Vardiya-Fazla Mesai Kanıt Sistemi"
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
    {"durum_id": "TASLAK", "durum_adi": "Taslak", "sira": 1,
     "izinli_sonraki_durumlar": "KONTROL|RED", "kategori": "acik"},
    {"durum_id": "KONTROL", "durum_adi": "Kontrol", "sira": 2,
     "izinli_sonraki_durumlar": "ONAY|TASLAK|RED", "kategori": "acik"},
    {"durum_id": "ONAY", "durum_adi": "Onay", "sira": 3,
     "izinli_sonraki_durumlar": "ARSIV", "kategori": "acik"},
    {"durum_id": "ARSIV", "durum_adi": "Arşiv", "sira": 4,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
    {"durum_id": "RED", "durum_adi": "Red", "sira": 5,
     "izinli_sonraki_durumlar": "", "kategori": "kapali"},
])

IZINLI_GECIS = [
    "TASLAK|KONTROL",
    "TASLAK|RED",
    "KONTROL|ONAY",
    "KONTROL|TASLAK",
    "KONTROL|RED",
    "ONAY|ARSIV",
]

PERSONEL = [
    ("PER-00001", "Ayşe Yılmaz", "Üretim", "Operatör"),
    ("PER-00002", "Mehmet Demir", "Üretim", "Usta"),
    ("PER-00003", "Zeynep Kaya", "Lojistik", "Depocu"),
    ("PER-00004", "Ali Çelik", "Lojistik", "Şoför"),
    ("PER-00005", "Elif Arslan", "Satış", "Temsilci"),
    ("PER-00006", "Can Öztürk", "Bakım", "Teknisyen"),
    ("PER-00007", "Selin Aydın", "İdari", "Asistan"),
    ("PER-00008", "Burak Şahin", "Üretim", "Operatör"),
]

IZIN_TURLERI = ["YOK", "YILLIK", "MAZERET", "RAPOR", "UCRETSIZ"]


def _yas_kova_pvf(yas_hucre: str) -> str:
    return (
        f'=IF({yas_hucre}="","",'
        f'IF({yas_hucre}<=pvf_kovaEsik1,"0-7",'
        f'IF({yas_hucre}<=pvf_kovaEsik2,"8-30",'
        f'IF({yas_hucre}<=pvf_kovaEsik3,"31-60","60+"))))'
    )


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    from openpyxl.comments import Comment
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Puantaj ve fazla mesai kararını etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı karar üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def _demo_akis():
    satirlar = []
    durumlar = ["TASLAK", "KONTROL", "ONAY", "ARSIV", "RED",
                "KONTROL", "ONAY", "TASLAK"]
    oncekiler = ["", "TASLAK", "KONTROL", "ONAY", "TASLAK",
                 "TASLAK", "KONTROL", "TASLAK"]
    for i in range(1, DEMO_AKIS + 1):
        kart = PERSONEL[(i - 1) % DEMO_KART][0]
        if i in (22, 25):
            kart = "" if i == 22 else "PER-99999"
        d = durumlar[(i - 1) % len(durumlar)]
        o = oncekiler[(i - 1) % len(oncekiler)]
        if i == 18:
            o, d = "TASLAK", "ONAY"
        if i == 20:
            o, d = "ONAY", "TASLAK"
        planli = 8.0
        giris = 8.0
        cikis = 18.0 if i % 3 == 0 else (17.0 if i % 3 == 1 else 16.5)
        fiili = cikis - giris
        hesap_fm = max(0.0, fiili - planli)
        bildir_fm = hesap_fm
        if i in (12, 19):
            bildir_fm = hesap_fm + 1.5
        if i == 15:
            bildir_fm = 4.5
            cikis = 20.0
            fiili = 12.0
            hesap_fm = 4.0
        tarih = RAPOR_TARIHI - timedelta(days=(i * 2) % 45)
        satirlar.append({
            "id": f"PUA-{i:05d}",
            "kart": kart,
            "tarih": tarih,
            "giris": giris,
            "cikis": cikis if i not in (12, 19) else (giris + planli + bildir_fm),
            "planli": planli,
            "fm": bildir_fm,
            "izin": IZIN_TURLERI[i % len(IZIN_TURLERI)],
            "onceki": o,
            "durum": d,
            "imza": f"İMZA-{i:03d}" if d in ("ONAY", "ARSIV") else "",
            "yorum_a": "A" if i % 2 else "B",
            "yorum_b": f"Puantaj notu {i}",
            "saat_ucret": 120 + (i % 5) * 10,
        })
    return satirlar


ORNEK = _demo_akis()


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Personel giriş/çıkış, planlı vardiya ve fazla mesaiyi yasal tavanla karşılaştırır; "
      "durum makinesi ve imza kanıt zinciri üretir; UYGUN / SINIR / AŞIM / VERİ YOK kararı verir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Makrosuz durum makinesi: Taslak → Kontrol → Onay → Arşiv / Red",
        "Günlük/haftalık fazla mesai + yıllık 270 saat tavan karşılaştırması",
        "Giriş/çıkış ↔ planlı vardiya ↔ bildirim FM zincir bayrağı",
        "Geçersiz geçiş ve yetim personel uyarıları",
        "KANIT_RAPORU — denetçi/SMMM imza alanlı dönem kanıtı",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "İK / bordro — puantaj kapanışı ve tavan kontrolü",
        "İşveren vekili — UYGUN / SINIR / AŞIM kararı",
        "Denetçi / SMMM — KANIT_RAPORU dosyalama",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) PERSONEL'e kart girin. 2) AKIS'e günlük puantaj satırlarını yazın. "
      "3) KUYRUKLAR ve PANO'dan kararı görün. 4) KANIT_RAPORU'ndan çıktı alın.",
      kaydir=True)
    ws.merge_cells("A19:L19")
    h(ws, 21, 1, f"Sürüm {SURUM} | Lisans: Tek kullanıcı | ExcelArşiv", yazi=GRİ, boyut=9)
    genislik(ws, {"A": 70})


def personel(ws):
    sayfa_hazirla(ws, "PERSONEL", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Personel kartları — her kaydın tek kimliği vardır", son_kolon=12)
    h(ws, 4, 1, "Sonraki Kimlik", yazi=GRİ)
    h(ws, 4, 2, "=pvf_sonrakiKimlik", kalin=True)
    yorum_ekle(ws, "B4", "Tanım: Otomatik kimlik önerisi | Neden önemli: ID disiplini | "
               "Doğru kullanım: Yeni kartta kullanın | Örnek: PER-00009 | Risk: Elle çakışma")

    sutunlar = ["PersonelId", "AdSoyad", "Departman", "Unvan", "HaftalikNormalSaat",
                "YillikMesaiTavan", "Aktif", "ToplamFm", "OrtGunlukFm"]
    baslik_satiri(ws, HDR, sutunlar)
    form = {
        "ToplamFm": (
            'IF(tblPersonel[[#This Row],[PersonelId]]="","",'
            'SUMIF(tblAkis[KaynakPersonelId],tblPersonel[[#This Row],[PersonelId]],'
            'tblAkis[FazlaMesaiSaat]))'
        ),
        "OrtGunlukFm": (
            'IF(tblPersonel[[#This Row],[PersonelId]]="","",'
            'IFERROR(AVERAGEIF(tblAkis[KaynakPersonelId],'
            'tblPersonel[[#This Row],[PersonelId]],tblAkis[FazlaMesaiSaat]),0))'
        ),
    }
    for i, (kid, ad, dep, unvan) in enumerate(PERSONEL):
        r = ILK + i
        _sari(ws, r, 1, kid, baslik="Personel Kimliği", mesaj="PER-##### formatında")
        _sari(ws, r, 2, ad, baslik="Ad Soyad", mesaj="Personel adı")
        _sari(ws, r, 3, dep, baslik="Departman", mesaj="Departman")
        _sari(ws, r, 4, unvan, baslik="Unvan", mesaj="Unvan")
        _sari(ws, r, 5, 45, sayi=CATI, baslik="Haftalık Normal", mesaj="Haftalık normal saat")
        _sari(ws, r, 6, 270, sayi=CATI, baslik="Yıllık Tavan", mesaj="Yıllık FM tavanı (saat)")
        _sari(ws, r, 7, "EVET", baslik="Aktif", mesaj="EVET veya HAYIR")
    for r in range(ILK + DEMO_KART, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=CATI if c in (5, 6) else None,
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblPersonel", f"A{HDR}:I{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Personel Kimliği", mesaj="PER-##### girin",
              hata_baslik="Kimlik", hata_mesaj="En az 1 karakter",
              isaret="greaterThan", f2="40")
    for c, ad in [(2, "Ad Soyad"), (3, "Departman"), (4, "Unvan")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "decimal", "0", f"E{ILK}:E{SON}",
              baslik="Haftalık Normal", mesaj="Saat ≥ 0",
              hata_baslik="Saat", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "decimal", "0", f"F{ILK}:F{SON}",
              baslik="Yıllık Tavan", mesaj="Saat ≥ 0",
              hata_baslik="Tavan", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeEvetHayir", f"G{ILK}:G{SON}",
              baslik="Aktif", mesaj="EVET veya HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 10)})
    ws.column_dimensions["B"].width = 22


def akis(ws):
    sayfa_hazirla(ws, "AKIS", RENK, URUN_AD, son_kolon=26)
    alt_bant(ws, 3,
             "Puantaj satırları — durum makinesi + yetim + geçersiz geçiş + zincir",
             son_kolon=20)
    sutunlar = [
        "PuantajId", "KaynakPersonelId", "Tarih", "GirisSaat", "CikisSaat",
        "PlanliVardiyaSaat", "FazlaMesaiSaat", "IzinTuru", "OncekiDurum", "Durum",
        "ImzaKanit", "YorumA", "YorumB", "SaatUcreti",
        "FiiliSaat", "HesaplananFm", "GecikmeGun", "YasKova",
        "YetimBayrak", "GecisUyarisi", "ZincirBayrak", "KayitDolu", "FmMaliyet",
    ]
    baslik_satiri(ws, HDR, sutunlar)

    yas_f = yaslandirma_gun(
        "tblAkis[[#This Row],[Tarih]]", "raporTarihi").lstrip("=")
    kova_f = _yas_kova_pvf("tblAkis[[#This Row],[GecikmeGun]]").lstrip("=")
    yetim_f = yetim_kayit_formul(
        "tblAkis[[#This Row],[KaynakPersonelId]]",
        "tblPersonel[PersonelId]").lstrip("=")
    gecis_birlesik = (
        'tblAkis[[#This Row],[OncekiDurum]]&"|"&tblAkis[[#This Row],[Durum]]'
    )
    gecis_f = gecersiz_gecis_formul(gecis_birlesik, "ListeIzinliGecis").lstrip("=")
    zincir_f = zincir_kirik_formul(
        "tblAkis[[#This Row],[HesaplananFm]]",
        "tblAkis[[#This Row],[FazlaMesaiSaat]]",
        "zincirTolerans",
    ).lstrip("=")

    form = {
        "FiiliSaat": (
            'IF(tblAkis[[#This Row],[PuantajId]]="","",'
            'MAX(0,tblAkis[[#This Row],[CikisSaat]]-tblAkis[[#This Row],[GirisSaat]]))'
        ),
        "HesaplananFm": (
            'IF(tblAkis[[#This Row],[PuantajId]]="","",'
            'MAX(0,tblAkis[[#This Row],[FiiliSaat]]-'
            'tblAkis[[#This Row],[PlanliVardiyaSaat]]))'
        ),
        "GecikmeGun": (
            f'IF(tblAkis[[#This Row],[PuantajId]]="","",'
            f'IF(OR(tblAkis[[#This Row],[Durum]]="ARSIV",'
            f'tblAkis[[#This Row],[Durum]]="RED"),0,{yas_f}))'
        ),
        "YasKova": f'IF(tblAkis[[#This Row],[PuantajId]]="","",{kova_f})',
        "YetimBayrak": f'IF(tblAkis[[#This Row],[PuantajId]]="","",{yetim_f})',
        "GecisUyarisi": (
            f'IF(OR(tblAkis[[#This Row],[PuantajId]]="",'
            f'tblAkis[[#This Row],[OncekiDurum]]=""),"",{gecis_f})'
        ),
        "ZincirBayrak": f'IF(tblAkis[[#This Row],[PuantajId]]="","",{zincir_f})',
        "KayitDolu": 'IF(tblAkis[[#This Row],[PuantajId]]="","",1)',
        "FmMaliyet": (
            'IF(tblAkis[[#This Row],[PuantajId]]="","",'
            'tblAkis[[#This Row],[FazlaMesaiSaat]]*'
            'tblAkis[[#This Row],[SaatUcreti]]*pvf_mesaiKatsayi)'
        ),
    }

    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["id"], baslik="Puantaj Kimliği", mesaj="PUA-#####")
        _sari(ws, r, 2, s["kart"] or None, baslik="Kaynak Personel",
              mesaj="PERSONEL'deki PersonelId")
        _sari(ws, r, 3, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Puantaj tarihi")
        _sari(ws, r, 4, s["giris"], sayi=CATI, baslik="Giriş Saat",
              mesaj="Giriş saati (ondalık, örn. 8)")
        _sari(ws, r, 5, s["cikis"], sayi=CATI, baslik="Çıkış Saat",
              mesaj="Çıkış saati (ondalık, örn. 17.5)")
        _sari(ws, r, 6, s["planli"], sayi=CATI, baslik="Planlı Vardiya",
              mesaj="Planlı vardiya saati")
        _sari(ws, r, 7, s["fm"], sayi=CATI, baslik="Fazla Mesai",
              mesaj="Bildirilen fazla mesai saati")
        _sari(ws, r, 8, s["izin"], baslik="İzin Türü", mesaj="İzin listesinden")
        _sari(ws, r, 9, s["onceki"] or None, baslik="Önceki Durum", mesaj="Önceki durum_id")
        _sari(ws, r, 10, s["durum"], baslik="Durum", mesaj="Durum haritasından seçin")
        _sari(ws, r, 11, s["imza"] or None, baslik="İmza/Kanıt", mesaj="İmza veya kanıt notu")
        _sari(ws, r, 12, s["yorum_a"], baslik="Yorum A", mesaj="A veya B")
        _sari(ws, r, 13, s["yorum_b"], baslik="Yorum B", mesaj="Serbest not")
        _sari(ws, r, 14, s["saat_ucret"], sayi=TL, baslik="Saat Ücreti",
              mesaj="Brüt saat ücreti TL")

    for r in range(ILK + DEMO_AKIS, SON + 1):
        for c in range(1, 15):
            fmt = TARİH if c == 3 else (
                TL if c == 14 else (CATI if c in (4, 5, 6, 7) else None))
            _sari(ws, r, c, None, sayi=fmt, baslik=sutunlar[c - 1], mesaj="Manuel giriş")

    tablo_ekle(ws, "tblAkis", f"A{HDR}:W{SON}", sutunlar, formuller=form)

    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}",
              baslik="Puantaj Kimliği", mesaj="PUA-##### girin",
              hata_baslik="Kimlik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "textLength", "0", f"B{ILK}:B{SON}",
              baslik="Kaynak Personel", mesaj="Personel kimliği (boş = yetim riski)",
              hata_baslik="Personel", hata_mesaj="Geçersiz",
              isaret="greaterThanOrEqual", f2="40")
    dogrulama(ws, "date", "1", f"C{ILK}:C{SON}",
              baslik="Tarih", mesaj="Puantaj tarihi",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih",
              isaret="greaterThan")
    for c, ad in [(4, "Giriş"), (5, "Çıkış"), (6, "Planlı"), (7, "FM")]:
        dogrulama(ws, "decimal", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} ≥ 0",
                  hata_baslik=ad, hata_mesaj="0 veya üzeri",
                  isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeIzin", f"H{ILK}:H{SON}",
              baslik="İzin Türü", mesaj="Listeden seçin",
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
              baslik="Saat Ücreti", mesaj="TL ≥ 0",
              hata_baslik="Ücret", hata_mesaj="0 veya üzeri",
              isaret="greaterThanOrEqual")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 12 for i in range(1, 24)})
    ws.column_dimensions["M"].width = 16
    ws.column_dimensions["T"].width = 16


def kuyruklar(ws):
    sayfa_hazirla(ws, "KUYRUKLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Durum bazlı canlı kuyruk — kapalı kayıtlar düşer", son_kolon=10)
    h(ws, 5, 1, "Kuyruk", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Adet", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "FM Saat", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    h(ws, 6, 1, "Kuyruk Taslak")
    h(ws, 6, 2, '=COUNTIFS(tblAkis[Durum],"TASLAK",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 6, 3, '=SUMIF(tblAkis[Durum],"TASLAK",tblAkis[FazlaMesaiSaat])', sayi=CATI)

    h(ws, 7, 1, "Kuyruk Kontrol")
    h(ws, 7, 2, '=COUNTIFS(tblAkis[Durum],"KONTROL",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 7, 3, '=SUMIF(tblAkis[Durum],"KONTROL",tblAkis[FazlaMesaiSaat])', sayi=CATI)

    h(ws, 8, 1, "Kuyruk Onay")
    h(ws, 8, 2, '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[KayitDolu],">0")', sayi=CATI)
    h(ws, 8, 3, '=SUMIF(tblAkis[Durum],"ONAY",tblAkis[FazlaMesaiSaat])', sayi=CATI)

    h(ws, 9, 1, "Aşım bandı (FM>günlük eşik)")
    h(ws, 9, 2,
      '=COUNTIFS(tblAkis[FazlaMesaiSaat],">"&pvf_esikGunlukFm,tblAkis[Durum],"<>ARSIV",'
      'tblAkis[Durum],"<>RED",tblAkis[KayitDolu],">0")',
      sayi=CATI)
    h(ws, 9, 3,
      '=SUMIFS(tblAkis[FazlaMesaiSaat],tblAkis[FazlaMesaiSaat],">"&pvf_esikGunlukFm,'
      'tblAkis[Durum],"<>ARSIV",tblAkis[Durum],"<>RED")',
      sayi=CATI)

    h(ws, 11, 1, "Kapalı (Arşiv+Red) — kuyruktan düşer", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 1, "Arşiv")
    h(ws, 12, 2, '=COUNTIF(tblAkis[Durum],"ARSIV")', sayi=CATI)
    h(ws, 12, 3, '=SUMIF(tblAkis[Durum],"ARSIV",tblAkis[FazlaMesaiSaat])', sayi=CATI)
    h(ws, 13, 1, "Red")
    h(ws, 13, 2, '=COUNTIF(tblAkis[Durum],"RED")', sayi=CATI)
    h(ws, 13, 3, '=SUMIF(tblAkis[Durum],"RED",tblAkis[FazlaMesaiSaat])', sayi=CATI)

    h(ws, 15, 1, "Kuyruk Özeti", kalin=True)
    h(ws, 15, 2, "=pvf_kuyrukOzeti", kaydir=True)
    ws.merge_cells("B15:F15")

    h(ws, 17, 1, "Yaş Kova Dağılımı", kalin=True, yazi=KOYU_LACIVERT)
    for i, kova in enumerate(["0-7", "8-30", "31-60", "60+"], 18):
        h(ws, i, 1, kova)
        h(ws, i, 2, f'=COUNTIF(tblAkis[YasKova],"{kova}")', sayi=CATI)
        h(ws, i, 3, f'=SUMIF(tblAkis[YasKova],"{kova}",tblAkis[FazlaMesaiSaat])', sayi=CATI)

    genislik(ws, {"A": 36, "B": 14, "C": 16})


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Hesap motoru — ≥40 adım; sabitler AYARLAR'dan", son_kolon=10)
    h(ws, 5, 1, "Adım", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Açıklama", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")

    adimlar = [
        ("H01", "Toplam puantaj", '=COUNTA(tblAkis[PuantajId])'),
        ("H02", "Dolu kayıt", '=SUM(tblAkis[KayitDolu])'),
        ("H03", "Toplam FM saat", "=SUM(tblAkis[FazlaMesaiSaat])"),
        ("H04", "Hesaplanan FM toplam", "=SUM(tblAkis[HesaplananFm])"),
        ("H05", "Fiili saat toplam", "=SUM(tblAkis[FiiliSaat])"),
        ("H06", "Taslak adet", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        ("H07", "Kontrol adet", '=COUNTIF(tblAkis[Durum],"KONTROL")'),
        ("H08", "Onay adet", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        ("H09", "Arşiv adet", '=COUNTIF(tblAkis[Durum],"ARSIV")'),
        ("H10", "Red adet", '=COUNTIF(tblAkis[Durum],"RED")'),
        ("H11", "Açık kuyruk adet", "=C11+C12+C13"),
        ("H12", "Kapalı düşüm adet", "=C14+C15"),
        ("H13", "Açık FM kuyruk",
         '=SUMIF(tblAkis[Durum],"TASLAK",tblAkis[FazlaMesaiSaat])'
         '+SUMIF(tblAkis[Durum],"KONTROL",tblAkis[FazlaMesaiSaat])'
         '+SUMIF(tblAkis[Durum],"ONAY",tblAkis[FazlaMesaiSaat])'),
        ("H14", "Yetim adet", '=COUNTIF(tblAkis[YetimBayrak],"YETİM")'),
        ("H15", "Geçersiz geçiş", '=COUNTIF(tblAkis[GecisUyarisi],"GEÇERSİZ GEÇİŞ")'),
        ("H16", "Zincir kırık", '=COUNTIF(tblAkis[ZincirBayrak],"ZİNCİR KIRIK")'),
        ("H17", "Zincir OK", '=COUNTIF(tblAkis[ZincirBayrak],"OK")'),
        ("H18", "Yaş 0-7", '=COUNTIF(tblAkis[YasKova],"0-7")'),
        ("H19", "Yaş 8-30", '=COUNTIF(tblAkis[YasKova],"8-30")'),
        ("H20", "Yaş 31-60", '=COUNTIF(tblAkis[YasKova],"31-60")'),
        ("H21", "Yaş 60+", '=COUNTIF(tblAkis[YasKova],"60+")'),
        ("H22", "Ort FM saat",
         "=IFERROR(AVERAGEIF(tblAkis[KayitDolu],1,tblAkis[FazlaMesaiSaat]),0)"),
        ("H23", "Günlük aşım adet",
         '=COUNTIFS(tblAkis[FazlaMesaiSaat],">"&pvf_esikGunlukFm,tblAkis[Durum],"<>ARSIV",'
         'tblAkis[Durum],"<>RED",tblAkis[KayitDolu],">0")'),
        ("H24", "Günlük aşım FM",
         '=SUMIFS(tblAkis[FazlaMesaiSaat],tblAkis[FazlaMesaiSaat],">"&pvf_esikGunlukFm,'
         'tblAkis[Durum],"<>ARSIV",tblAkis[Durum],"<>RED")'),
        ("H25", "FM maliyet toplam",
         "=IFERROR(SUMIF(tblAkis[KayitDolu],1,tblAkis[FmMaliyet]),0)"),
        ("H26", "Yıllık tavan aşım (proxy)",
         '=IF(C8>pvf_yillikTavan,C8-pvf_yillikTavan,0)'),
        ("H27", "Haftalık FM proxy", "=C8*pvf_haftalikCarpan"),
        ("H28", "FM sapma (bildirim-hesap)", "=C8-C9"),
        ("H29", "Temkinli senaryo", "=C18*pvf_senaryoTemkinli"),
        ("H30", "Baz senaryo", "=C18*pvf_senaryoBaz"),
        ("H31", "İyimser senaryo", "=C18*pvf_senaryoIyimser"),
        ("H32", "Senaryo farkı", "=C36-C34"),
        ("H33", "Tornado (oran×FM)", "=C18*pvf_tornadoOran"),
        ("H34", "HHI proxy",
         "=IFERROR((MAX(C11,C12,C13)/MAX(C16,1))^2,0)"),
        ("H35", "Kalite skoru",
         "=IF(C6=0,0,MAX(0,pvf_kaliteTaban-C19*pvf_kaliteCeza-C20*pvf_kaliteCeza"
         "-C21*pvf_kaliteCeza-C28*pvf_kaliteSapma))"),
        ("H36", "Tahmin PERCENTILE",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[FazlaMesaiSaat],pvf_yuzdelikOran),0)"),
        ("H37", "Tahmin FORECAST",
         "=IFERROR(FORECAST(COUNTA(tblAkis[PuantajId])+1,tblAkis[FazlaMesaiSaat],"
         "tblAkis[KayitDolu]),C41)"),
        ("H38", "Tahmin aralık", "=IFERROR((C41+C42)/2,0)"),
        ("H39", "P90 yüzdelik",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblAkis[FazlaMesaiSaat],pvf_yuzdelikOran),0)"),
        ("H40", "Personel sayısı", '=COUNTA(tblPersonel[PersonelId])'),
        ("H41", "Tavan toplam", "=SUM(tblPersonel[YillikMesaiTavan])"),
        ("H42", "Tavan kullanım oranı", "=IFERROR(C8/MAX(C46,1),0)"),
        ("H43", "Kuyruk özeti metin",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Taslak",C11,"Kontrol",C12,"Onay",C13,"Aşım",C28)'),
        ("H44", "Yaş kova metin",
         '=_xlfn.TEXTJOIN("/",TRUE,C23,C24,C25,C26)'),
        ("H45", "Senaryo motor özet",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Temkinli",C34,"Baz",C35,"İyimser",C36)'),
        ("H46", "Veri hazırlık", "=IF(C6=0,0,C7/MAX(C6,1))"),
        ("H47", "Sınır adet (tavan %90)",
         '=IF(AND(C8>=pvf_yillikTavan*pvf_sinirOran,C8<=pvf_yillikTavan),1,0)'),
        ("H48", "Aşım karar adet",
         '=IF(OR(C28>0,C31>0),1,0)'),
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
      '=IF(COUNTA(tblAkis[PuantajId])=0,"VERİ YOK",'
      'IF(pvf_asimAdet>0,"AŞIM",'
      'IF(OR(pvf_sinirAdet>0,pvf_gunlukAsim>0),"SINIR","UYGUN")))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Açık Kuyruk Adet", "=pvf_acikAdet", CATI),
        (7, "Taslak Adet", '=COUNTIFS(tblAkis[Durum],"TASLAK",tblAkis[KayitDolu],">0")', CATI),
        (8, "Kontrol Adet", '=COUNTIFS(tblAkis[Durum],"KONTROL",tblAkis[KayitDolu],">0")', CATI),
        (9, "Onay Adet", '=COUNTIFS(tblAkis[Durum],"ONAY",tblAkis[KayitDolu],">0")', CATI),
        (10, "Toplam FM Saat", "=pvf_toplamFm", CATI),
        (11, "FM Maliyet", "=pvf_fmMaliyet", TL),
        (12, "Yetim Kayıt", "=pvf_yetimKayit", CATI),
        (13, "Geçersiz Geçiş", "=pvf_gecersizGecis", CATI),
        (14, "Zincir Kırık", "=pvf_zincirKirik", CATI),
        (15, "Yaş Kova Özeti", "=pvf_yasKova", None),
        (16, "Günlük FM Ort.", "=pvf_gunlukFmOrt", CATI),
        (17, "Tahmin Aralık", "=pvf_tahminAralik", CATI),
        (18, "Senaryo Farkı", "=pvf_senaryoKarsilastirma", CATI),
        (19, "Tornado Zirve", "=pvf_tornadoZirve", CATI),
        (20, "HHI Yoğunlaşma", "=pvf_hhi", YÜZDE),
        (21, "Kalite Skoru", "=pvf_kaliteSkor", CATI),
        (22, "P90 Yüzdelik", "=pvf_yuzdelikP90", CATI),
        (23, "Kapalı Düşüm", "=pvf_kapaliDusum", CATI),
        (24, "Haftalık FM", "=pvf_haftalikFm", CATI),
        (25, "FM Sapma", "=pvf_fmSapma", CATI),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    for r, form in [
        (7, "=pvf_yorumGunluk"),
        (8, "=pvf_yorumHaftalik"),
        (9, "=pvf_yorumSapma"),
        (10, "=pvf_yorumTornado"),
        (11, "=pvf_yorumSenaryo"),
        (12, "=pvf_yorumHhi"),
        (13, "=pvf_yorumKalite"),
        (14, "=pvf_yorumSenMotor"),
        (15, "=pvf_yorumTahmin"),
        (16, "=pvf_yorumP90"),
    ]:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 27, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 1, "Durum")
    h(ws, 28, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Durum Taslak", 6), ("Durum Kontrol", 8), ("Durum Onay", 5),
        ("Durum Arşiv", 5), ("Durum Red", 4),
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
    h(ws, 28, 8, "FM Saat")
    for i in range(8):
        h(ws, 29 + i, 7, i + 1, sayi=CATI)
        h(ws, 29 + i, 8, 12 + i * 2, sayi=CATI)

    h(ws, 28, 10, "Senaryo")
    h(ws, 28, 11, "FM Saat")
    for i, (ad, t) in enumerate([
        ("Temkinli", 38), ("Baz", 45), ("İyimser", 52),
    ], 29):
        h(ws, i, 10, ad)
        h(ws, i, 11, t, sayi=CATI)

    h(ws, 28, 13, "Departman")
    h(ws, 28, 14, "FM")
    for i, (ad, t) in enumerate([
        ("Üretim", 22), ("Lojistik", 14), ("Satış", 8), ("Bakım", 6),
    ], 29):
        h(ws, i, 13, ad)
        h(ws, i, 14, t, sayi=CATI)

    c1 = PieChart()
    c1.title = "Durum Dağılımı"
    c1.add_data(Reference(ws, min_col=2, min_row=28, max_row=33), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=29, max_row=33))
    c1.width, c1.height = 10, 7
    ws.add_chart(c1, "A42")

    c2 = BarChart()
    c2.title = "Yaşlandırma Kovaları"
    c2.add_data(Reference(ws, min_col=5, min_row=28, max_row=32), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=29, max_row=32))
    c2.width, c2.height = 10, 7
    ws.add_chart(c2, "F42")

    c3 = LineChart()
    c3.title = "Haftalık Fazla Mesai"
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
    c5.title = "Departman FM Payı"
    c5.add_data(Reference(ws, min_col=14, min_row=28, max_row=32), titles_from_data=True)
    c5.set_categories(Reference(ws, min_col=13, min_row=29, max_row=32))
    c5.width, c5.height = 10, 7
    ws.add_chart(c5, "A72")

    c6 = BarChart()
    c6.title = "Kuyruk FM Saat"
    c6.add_data(Reference(ws, min_col=2, min_row=28, max_row=31), titles_from_data=True)
    c6.set_categories(Reference(ws, min_col=1, min_row=29, max_row=31))
    c6.width, c6.height = 10, 7
    ws.add_chart(c6, "F72")

    c7 = LineChart()
    c7.title = "FM Trend (örnek)"
    c7.add_data(Reference(ws, min_col=8, min_row=28, max_row=34), titles_from_data=True)
    c7.set_categories(Reference(ws, min_col=7, min_row=29, max_row=34))
    c7.width, c7.height = 12, 7
    ws.add_chart(c7, "A87")

    c8 = BarChart()
    c8.title = "Senaryo vs Departman"
    c8.add_data(Reference(ws, min_col=14, min_row=28, max_row=31), titles_from_data=True)
    c8.set_categories(Reference(ws, min_col=13, min_row=29, max_row=31))
    c8.width, c8.height = 10, 7
    ws.add_chart(c8, "F87")

    baski_hazirla(ws, "A1:N40", f"{URUN_AD} | Pano | {SURUM}")
    genislik(ws, {"A": 22, "B": 16, "C": 14, "D": 14, "E": 12})


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "PVF-001..005 kural tablosu — 2025 / 2026 / 2027", son_kolon=12)
    kurallar_data = [
        ("PVF-001", "4857/63", "Haftalık normal çalışma süresi", 45, 45, 45, "saat"),
        ("PVF-002", "4857/41", "Yıllık fazla çalışma tavanı", 270, 270, 270, "saat"),
        ("PVF-003", "4857/41", "Fazla çalışma ücret katsayısı", 1.5, 1.5, 1.5, "oran"),
        ("PVF-004", "4857/41", "Fazla sürelerle çalışma katsayısı", 1.25, 1.25, 1.25, "oran"),
        ("PVF-005", "İşyeri", "İmza/kanıt zorunluluğu (ONAY/ARSİV)", 1, 1, 1, "evet"),
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
            'IF(pvf_aktifYil=2025,tblKurallar[[#This Row],[y2025]],'
            'IF(pvf_aktifYil=2027,tblKurallar[[#This Row],[y2027]],'
            'tblKurallar[[#This Row],[y2026]]))'
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A5:H10", basliklar, formuller=form_k)
    h(ws, 12, 1, "Aktif yıl", kalin=True)
    _sari(ws, 12, 2, 2026, sayi=CATI, baslik="Aktif yıl", mesaj="2025/2026/2027")
    h(ws, 13, 1, "Kural yıl matrisi", kalin=True)
    h(ws, 13, 2, "=pvf_kuralYilMatris", kaydir=True)
    ws.merge_cells("B13:F13")
    genislik(ws, {"A": 12, "B": 12, "C": 42, "D": 10, "E": 10, "F": 10, "G": 10})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Örnek vakalar — beklenen karar ve fark kontrolü", son_kolon=10)
    basliklar = ["vaka_id", "aciklama", "fm_saat", "tavan", "beklenen_karar", "fark"]
    baslik_satiri(ws, 5, basliklar)
    vakalar_data = [
        ("V01", "Normal vardiya, FM yok", 0, 270, "UYGUN"),
        ("V02", "Günlük 2 saat FM", 2, 270, "UYGUN"),
        ("V03", "Günlük eşik üstü FM", 4.5, 270, "SINIR"),
        ("V04", "Yıllık tavan aşımı", 280, 270, "AŞIM"),
        ("V05", "Boş veri seti", 0, 270, "VERİ YOK"),
        ("V06", "Zincir sapmalı bildirim", 3, 270, "SINIR"),
    ]
    for i, (vid, acik, fm, tavan, karar) in enumerate(vakalar_data):
        r = 6 + i
        _sari(ws, r, 1, vid, baslik="Vaka", mesaj="Vaka kimliği")
        _sari(ws, r, 2, acik, baslik="Açıklama", mesaj="Vaka açıklaması")
        _sari(ws, r, 3, fm, sayi=CATI, baslik="FM", mesaj="FM saat")
        _sari(ws, r, 4, tavan, sayi=CATI, baslik="Tavan", mesaj="Tavan saat")
        _sari(ws, r, 5, karar, baslik="Beklenen", mesaj="Beklenen karar")
    form_v = {
        "fark": (
            'IF(tblVakalar[[#This Row],[vaka_id]]="","",'
            'IF(tblVakalar[[#This Row],[fm_saat]]>tblVakalar[[#This Row],[tavan]],"AŞIM",'
            'IF(tblVakalar[[#This Row],[fm_saat]]>pvf_esikGunlukFm,"SINIR",'
            'IF(AND(tblVakalar[[#This Row],[fm_saat]]=0,'
            'tblVakalar[[#This Row],[vaka_id]]="V05"),"VERİ YOK","UYGUN"))))'
        ),
    }
    tablo_ekle(ws, "tblVakalar", "A5:F11", basliklar, formuller=form_v)
    h(ws, 13, 1, "Vaka fark özeti", kalin=True)
    h(ws, 13, 2,
      '=COUNTIF(tblVakalar[fark],"<>"&"")&" satır"', kaydir=True)
    genislik(ws, {"A": 10, "B": 32, "C": 12, "D": 10, "E": 16, "F": 14})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Puantaj / Fazla Mesai Kanıt Raporu", boyut=16, son_kolon=10)
    for i, (etiket, deger) in enumerate(rapor_baslik_satirlari(URUN_AD, SURUM), 4):
        h(ws, i, 1, etiket)
        h(ws, i, 2, deger)
    h(ws, 8, 1, "Rapor Tarihi")
    h(ws, 8, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 9, 1, "KARAR", kalin=True)
    h(ws, 9, 2, "=PANO!B4", kalin=True, boyut=14)

    ozet = [
        (11, "Açık Kuyruk", "=pvf_acikAdet", CATI),
        (12, "Toplam FM Saat", "=pvf_toplamFm", CATI),
        (13, "FM Maliyet", "=pvf_fmMaliyet", TL),
        (14, "Yetim", "=pvf_yetimKayit", CATI),
        (15, "Geçersiz Geçiş", "=pvf_gecersizGecis", CATI),
        (16, "Zincir Kırık", "=pvf_zincirKirik", CATI),
        (17, "Kural Matris", "=pvf_kuralYilMatris", None),
        (18, "Kanıt Özeti", "=pvf_kanitRaporu", None),
        (19, "Haftalık FM", "=pvf_haftalikFm", CATI),
        (20, "Kalite Skoru", "=pvf_kaliteSkor", CATI),
        (21, "Kapalı Düşüm", "=pvf_kapaliDusum", CATI),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)

    h(ws, 23, 1, "Hesap Zinciri", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 1,
      "Giriş/Çıkış → Fiili saat → Planlı vardiya farkı → Hesaplanan FM → "
      "Bildirilen FM (zincir) → Durum makinesi → PANO kararı → İmza/kanıt.",
      kaydir=True)
    ws.merge_cells("A24:H24")

    h(ws, 26, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 1,
      "Yıllık tavan PVF-002 (270 saat) AYARLAR'dan gelir. Geçersiz geçiş engellenmez, "
      "görünür kılınır. Bu çıktı karar destektir; kesin hukuki/bordro görüşü yerine geçmez.",
      kaydir=True)
    ws.merge_cells("A27:H27")

    h(ws, 29, 1, "İmza / Onay", kalin=True, yazi=KOYU_LACIVERT)
    for i, (etiket, deger) in enumerate(imza_alani(), 30):
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
        (5, "Puantaj sayısı", '=COUNTA(tblAkis[PuantajId])'),
        (6, "Personel sayısı", '=COUNTA(tblPersonel[PersonelId])'),
        (7, "Taslak", '=COUNTIF(tblAkis[Durum],"TASLAK")'),
        (8, "Kontrol", '=COUNTIF(tblAkis[Durum],"KONTROL")'),
        (9, "Onay", '=COUNTIF(tblAkis[Durum],"ONAY")'),
        (10, "Arşiv", '=COUNTIF(tblAkis[Durum],"ARSIV")'),
        (11, "Red", '=COUNTIF(tblAkis[Durum],"RED")'),
        (12, "Açık toplam", "=B7+B8+B9"),
        (13, "Yetim", "=pvf_yetimKayit"),
        (14, "Geçersiz", "=pvf_gecersizGecis"),
        (15, "Zincir kırık", "=pvf_zincirKirik"),
        (16, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (17, "Kalite skoru", "=pvf_kaliteSkor"),
        (18, "Toplam FM", "=pvf_toplamFm"),
        (19, "FM maliyet", "=pvf_fmMaliyet"),
        (20, "Kapalı düşüm", "=pvf_kapaliDusum"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r == 19 else (CATI if r < 16 or r in (17, 18, 20) else None))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO_KART} personel kartı ve {DEMO_AKIS} puantaj satırı vardır. "
      "Yetim (22, 25), geçersiz geçiş (18, 20) ve zincir kırık (12, 19) kasıtlıdır. "
      "PERSONEL ve AKIS'i temizleyip kendi verinizi girebilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Durumlar: Taslak/Kontrol/Onay/Arşiv/Red — kapalılar kuyruktan düşer")
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

    h(ws, 3, 9, "İzin Türü", kalin=True)
    h(ws, 6, 9, "Izin", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(IZIN_TURLERI, 7):
        _sari(ws, i, 9, b, baslik="İzin", mesaj="İzin listesi")
    genislik(ws, {"A": 14, "C": 22, "E": 12, "G": 10, "I": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler + durum haritası (A01/A06)", son_kolon=12)

    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026",
         "Rapor tarihi (TODAY yok)"),
        ("zincirTolerans", 1, "saat", "İş kuralı — FM sapma", "01.01.2026",
         "ZİNCİR KIRIK eşiği (saat)"),
        ("pvf_esikGunlukFm", 3, "saat", "İç politika / PVF", "01.01.2026", "Günlük FM eşiği"),
        ("pvf_yillikTavan", 270, "saat", "4857 md.41 / PVF-002", "01.01.2026", "Yıllık FM tavan"),
        ("pvf_sinirOran", 0.9, "oran", "İç politika", "01.01.2026", "SINIR bandı oranı"),
        ("pvf_kovaEsik1", 7, "gün", "Yaşlandırma", "01.01.2026", "Kova 1 üst sınır"),
        ("pvf_kovaEsik2", 30, "gün", "Yaşlandırma", "01.01.2026", "Kova 2 üst sınır"),
        ("pvf_kovaEsik3", 60, "gün", "Yaşlandırma", "01.01.2026", "Kova 3 üst sınır"),
        ("pvf_mesaiKatsayi", 1.5, "oran", "4857 md.41 / PVF-003", "01.01.2026", "FM ücret katsayısı"),
        ("pvf_haftalikCarpan", 0.2, "oran", "Haftalık proxy", "01.01.2026", "Haftalık FM çarpan"),
        ("pvf_senaryoTemkinli", 0.85, "oran", "Senaryo motoru", "01.01.2026", "Temkinli çarpan"),
        ("pvf_senaryoBaz", 1.0, "oran", "Senaryo motoru", "01.01.2026", "Baz çarpan"),
        ("pvf_senaryoIyimser", 1.15, "oran", "Senaryo motoru", "01.01.2026", "İyimser çarpan"),
        ("pvf_tornadoOran", 0.08, "oran", "Duyarlılık", "01.01.2026", "Tornado etki oranı"),
        ("pvf_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("pvf_olcekHedef", 20000, "satir", "Manda A3 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("pvf_azamiFm", 12, "saat", "İç politika — uç değer", "01.01.2026", "Azami günlük FM"),
        ("pvf_kaliteTaban", 100, "puan", "Kalite modeli", "01.01.2026", "Kalite taban puanı"),
        ("pvf_kaliteCeza", 10, "puan", "Kalite modeli", "01.01.2026", "Yetim/geçiş/zincir cezası"),
        ("pvf_kaliteSapma", 2, "puan", "Kalite modeli", "01.01.2026", "Sapma cezası"),
        ("pvf_durumHaritasi", "AYARLAR durum tablosu", "metin", "SPEC akis.durum_haritasi",
         "01.01.2026", "Durum makinesi kaynağı"),
        ("pvf_kuyrukOzeti_metin", "KUYRUKLAR canlı özet", "metin", "A05 kuyruk", "01.01.2026",
         "Kuyruk görünümü"),
        ("pvf_kapaliDusum_metin", "kapali kategori kuyruktan düşer", "metin", "A06", "01.01.2026",
         "Kapalı düşüm kuralı"),
        ("pvf_kanitRaporu_metin", "KANIT_RAPORU çıktısı", "metin", "A4 kanıt", "01.01.2026",
         "Kanıt raporu"),
        ("pvf_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("pvf_firmaUnvan", "Örnek Üretim A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026",
         "Firma"),
        ("pvf_aktifYil", 2026, "yil", "KURALLAR", "01.01.2026", "Aktif hesap yılı"),
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
        (44, "pvf_acikAdet", "=HESAP!C16"),
        (45, "pvf_toplamFm", "=HESAP!C8"),
        (46, "pvf_gunlukAsim", "=HESAP!C28"),
        (47, "pvf_fmMaliyet", "=HESAP!C30"),
        (48, "pvf_yetimKayit", "=HESAP!C19"),
        (49, "pvf_gecersizGecis", "=HESAP!C20"),
        (50, "pvf_zincirKirik", "=HESAP!C21"),
        (51, "pvf_asimAdet", "=HESAP!C53"),
        (52, "pvf_yasKova", "=HESAP!C49"),
        (53, "pvf_gunlukFmOrt", "=HESAP!C27"),
        (54, "pvf_fmSapma", "=HESAP!C33"),
        (55, "pvf_tahminAralik", "=HESAP!C43"),
        (56, "pvf_senaryoKarsilastirma", "=HESAP!C37"),
        (57, "pvf_tornadoZirve", "=HESAP!C38"),
        (58, "pvf_hhi", "=HESAP!C39"),
        (59, "pvf_kaliteSkor", "=HESAP!C40"),
        (60, "pvf_yuzdelikP90", "=HESAP!C44"),
        (61, "pvf_senaryoMotor", "=HESAP!C50"),
        (62, "pvf_kapaliDusum", "=HESAP!C17"),
        (63, "pvf_kuyrukOzeti", "=HESAP!C48"),
        (64, "pvf_kanitRaporu",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"Açık",HESAP!C16,"Yetim",HESAP!C19,'
         '"Zincir",HESAP!C21,"FM",HESAP!C8,"Aşım",HESAP!C53)'),
        (65, "pvf_sonrakiKimlik",
         sonraki_kimlik_formul("PER", "tblPersonel[PersonelId]").replace(
             'MAX(tblPersonel[PersonelId])',
             'COUNTA(tblPersonel[PersonelId])')),
        (66, "pvf_haftalikFm", "=HESAP!C32"),
        (67, "pvf_sinirAdet", "=HESAP!C52"),
        (68, "pvf_kuralYilMatris",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"PVF-001..005","2025-2027",pvf_aktifYil)'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    h(ws, 70, 8, "Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (71, "pvf_yorumGunluk",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Günlük FM ort",TEXT(pvf_gunlukFmOrt,"0.0"),"saat")'),
        (72, "pvf_yorumHaftalik",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Haftalık FM proxy",TEXT(pvf_haftalikFm,"0.0"),"saat")'),
        (73, "pvf_yorumSapma",
         '=_xlfn.TEXTJOIN(" ",TRUE,"FM sapma",TEXT(pvf_fmSapma,"0.00"),"saat")'),
        (74, "pvf_yorumTornado",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tornado zirve",TEXT(pvf_tornadoZirve,"0.00"),"saat")'),
        (75, "pvf_yorumSenaryo",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo farkı",TEXT(pvf_senaryoKarsilastirma,"0.00"))'),
        (76, "pvf_yorumHhi",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Yoğunlaşma HHI",TEXT(pvf_hhi,"0.0%"))'),
        (77, "pvf_yorumKalite",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Kalite skoru",pvf_kaliteSkor)'),
        (78, "pvf_yorumSenMotor",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Senaryo motoru",pvf_senaryoMotor)'),
        (79, "pvf_yorumTahmin",
         '=_xlfn.TEXTJOIN(" ",TRUE,"Tahmin aralık",TEXT(pvf_tahminAralik,"0.00"),"saat")'),
        (80, "pvf_yorumP90",
         '=_xlfn.TEXTJOIN(" ",TRUE,"P90 FM",TEXT(pvf_yuzdelikP90,"0.00"),"saat")'),
    ]
    for r, ad, form in yorumlar:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    genislik(ws, {"A": 24, "B": 28, "C": 10, "D": 28, "E": 14, "F": 28, "H": 26, "I": 55})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    maddeler = [
        "1. PERSONEL sayfasına personel kartı ekleyin; sonraki kimliği kullanın.",
        "2. AKIS sayfasına her günün giriş/çıkış, planlı vardiya ve FM satırını yazın.",
        "3. Önceki durum boş bırakılırsa geçiş uyarısı çalışmaz; bilinçli geçişte doldurun.",
        "4. Geçersiz geçiş engellenmez — görünür uyarı üretir (makrosuz kural).",
        "5. Kapalı durumlar (Arşiv/Red) açık kuyruklardan düşer.",
        "6. raporTarihi AYARLAR'dadır; TODAY kullanılmaz. Yaşlandırma puantaj tarihine göredir.",
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
    for r in range(6, 30):
        dogrulama(ws, "textLength", "0", f"D{r}",
                  baslik="Kaynak", mesaj="Kaynak metni girin",
                  hata_baslik="Kaynak", hata_mesaj="Boş bırakmayın",
                  isaret="greaterThanOrEqual", f2="120")
    ws = wb["LISTELER"]
    for r in range(7, 13):
        dogrulama(ws, "textLength", "1", f"A{r}",
                  baslik="Durum", mesaj="Durum kimliği",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    for r in range(7, 13):
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
                  baslik="Kural", mesaj="PVF-###",
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
        "B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"SINIR"'], fill=sari))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"AŞIM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 26):
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
    for durum, fill in [("KONTROL", sari), ("RED", kirmizi), ("ONAY", yesil)]:
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
        "B9", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"SINIR"'], fill=sari))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"AŞIM"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(
        "B9", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(11, 22):
        ws.conditional_formatting.add(
            f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))

    ws = wb["PERSONEL"]
    ws.conditional_formatting.add(
        f"G{ILK}:G{SON}",
        CellIsRule(operator="equal", formula=['"HAYIR"'], fill=sari))
    ws.conditional_formatting.add(
        f"F{ILK}:F{SON}",
        CellIsRule(operator="greaterThan", formula=["270"], fill=yesil))


def adlari_bagla(wb):
    param_adlari = [
        "raporTarihi", "zincirTolerans", "pvf_esikGunlukFm", "pvf_yillikTavan",
        "pvf_sinirOran", "pvf_kovaEsik1", "pvf_kovaEsik2", "pvf_kovaEsik3",
        "pvf_mesaiKatsayi", "pvf_haftalikCarpan",
        "pvf_senaryoTemkinli", "pvf_senaryoBaz", "pvf_senaryoIyimser",
        "pvf_tornadoOran", "pvf_yuzdelikOran", "pvf_olcekHedef", "pvf_azamiFm",
        "pvf_kaliteTaban", "pvf_kaliteCeza", "pvf_kaliteSapma",
        "pvf_durumHaritasi", "pvf_kuyrukOzeti_metin", "pvf_kapaliDusum_metin",
        "pvf_kanitRaporu_metin", "pvf_dosyaSurumu", "pvf_firmaUnvan", "pvf_aktifYil",
    ]
    for i, ana in enumerate(param_adlari):
        ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    motor_map = {
        "pvf_acikAdet": 44, "pvf_toplamFm": 45, "pvf_gunlukAsim": 46,
        "pvf_fmMaliyet": 47, "pvf_yetimKayit": 48, "pvf_gecersizGecis": 49,
        "pvf_zincirKirik": 50, "pvf_asimAdet": 51, "pvf_yasKova": 52,
        "pvf_gunlukFmOrt": 53, "pvf_fmSapma": 54, "pvf_tahminAralik": 55,
        "pvf_senaryoKarsilastirma": 56, "pvf_tornadoZirve": 57, "pvf_hhi": 58,
        "pvf_kaliteSkor": 59, "pvf_yuzdelikP90": 60, "pvf_senaryoMotor": 61,
        "pvf_kapaliDusum": 62, "pvf_kuyrukOzeti": 63, "pvf_kanitRaporu": 64,
        "pvf_sonrakiKimlik": 65, "pvf_haftalikFm": 66, "pvf_sinirAdet": 67,
        "pvf_kuralYilMatris": 68,
    }
    for ad, r in motor_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    yorum_map = {
        "pvf_yorumGunluk": 71, "pvf_yorumHaftalik": 72, "pvf_yorumSapma": 73,
        "pvf_yorumTornado": 74, "pvf_yorumSenaryo": 75, "pvf_yorumHhi": 76,
        "pvf_yorumKalite": 77, "pvf_yorumSenMotor": 78, "pvf_yorumTahmin": 79,
        "pvf_yorumP90": 80,
    }
    for ad, r in yorum_map.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeDurumlar", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIzinliGecis", "LISTELER!$C$7:$C$12")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")
    ad_ekle(wb, "ListeYorumAB", "LISTELER!$G$7:$G$8")
    ad_ekle(wb, "ListeIzin", "LISTELER!$I$7:$I$11")


def main(cikti_yolu=None):
    wb = Workbook()
    siralar = [
        (kapak, "KAPAK"),
        (personel, "PERSONEL"),
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
    dosya_ad = "PuantajVardiyaFazlaMesaiKanitSistemi.xlsx"
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
