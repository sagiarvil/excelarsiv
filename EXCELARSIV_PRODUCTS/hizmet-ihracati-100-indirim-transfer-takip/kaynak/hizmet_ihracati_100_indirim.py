#!/usr/bin/env python3
"""Hizmet İhracatı %100 İndirim + Transfer Takip — A1+A2 üretim betiği (manda v6)."""

from __future__ import annotations

import os
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if KOK not in sys.path:
    sys.path.insert(0, KOK)

from excel_uretim.ortak import (  # noqa: E402
    CATI,
    DIS_REF_YEŞIL,
    FONT,
    GIRIS_SARI,
    GIRIS_YAZI,
    GRİ,
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
from ortak.mevzuat_motoru import (  # noqa: E402
    KURAL_BASLIK,
    MOTOR_BASLIK,
    VAKA_BASLIK,
)
from ortak.mutabakat_motoru import (  # noqa: E402
    dort_kova_butunluk,
    itiraz_metni_sablon,
    kok_neden_onerisi,
    normalize_formul,
)

URUN_AD = "Hizmet İhracatı %100 İndirim + Transfer Takip"
SURUM = "1.0.0"
RENK = "1F4E79"
KAPASITE = 500
DEMO = 35
HDR = 5
ILK = HDR + 1
SON = HDR + KAPASITE
RAPOR_TARIHI = date(2026, 8, 10)

ALICI_ULKELER = ["Almanya", "Hollanda", "Birleşik Krallık", "ABD", "Diğer"]
DURUMLAR = ["Açık", "Kapalı", "İtiraz", "Beklemede"]
BANKALAR = ["Ziraat", "İş Bankası", "Garanti", "Yapı Kredi", "Diğer"]


def _ornek_satirlar():
    satirlar = []
    for i in range(1, DEMO + 1):
        no = f"HIH{100000 + i}"
        if i in (3, 7, 11):
            no_a, no_b = no[:10], no
        elif i == 15:
            no_a, no_b = no + "X", no
        else:
            no_a, no_b = no, no
        fatura = 5000 + i * 370
        stopaj = round(fatura * 0.02, 2) if i % 6 == 0 else 0
        kur = round(fatura * 0.005, 2) if i % 5 == 0 else 0
        kismi = 150 if i % 9 == 0 else 0
        diger = round(fatura * 0.01, 2) if i % 7 == 0 else 0
        net = round(fatura - stopaj - kur - kismi - diger, 2)
        if i % 8 == 0:
            b_tutar = round(net - 15.5, 2)
        elif i % 11 == 0:
            b_tutar = round(net + 0.03, 2)
        elif i == 15:
            b_tutar = net
        else:
            b_tutar = net
        ulke = ALICI_ULKELER[i % len(ALICI_ULKELER)]
        tar = RAPOR_TARIHI - timedelta(days=(DEMO - i))
        satirlar.append({
            "no_a": no_a, "no_b": no_b, "tarih": tar, "fatura": fatura,
            "stopaj": stopaj, "kur": kur, "kismi": kismi, "diger": diger,
            "net": net, "b_tutar": b_tutar, "ulke": ulke,
            "banka": BANKALAR[i % len(BANKALAR)],
        })
    return satirlar


ORNEK = _ornek_satirlar()


def _sari(ws, r, c, deger, sayi=None, baslik="", mesaj=""):
    hucre = h(ws, r, c, deger, zemin=GIRIS_SARI, yazi=GIRIS_YAZI, sayi=sayi)
    if baslik:
        from openpyxl.comments import Comment
        hucre.comment = Comment(
            f"Tanım: {baslik} | Neden önemli: Mutabakat ve indirim sonucunu etkiler | "
            f"Doğru kullanım: {mesaj or 'Uygun türde değer girin'} | "
            f"Örnek: hücredeki örnek değere bakın | Risk: Yanlış giriş hatalı kova üretir",
            "ExcelArşiv", height=160, width=300)
    return hucre


def kural_cek(kural_id: str) -> str:
    return (
        f'=IFERROR(INDEX(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili="",hesapYili=0),1,'
        f'hesapYili-hih_yilTaban))),'
        f'tblKurallar[deger_2025],tblKurallar[deger_2026],tblKurallar[deger_2027]),'
        f'MATCH("{kural_id}",tblKurallar[kural_id],0)),0)'
    )


def _guvenli_formul(formul: str, birim: str) -> str:
    if not formul.startswith("="):
        formul = "=" + formul
    inner = formul[1:]
    if inner.upper().startswith("IFERROR("):
        return formul
    yedek = '""' if birim in ("metin", "kod") else "0"
    return f"=IFERROR({inner},{yedek})"


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=20)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=20)
    h(ws, 4, 1,
      "Hizmet ihracatı faturalarını banka transferleri ile tek anahtarda eşleştirir; "
      "dört kovaya ayırır; %100 indirim uygunluğunu ve kanıt zincirini üretir.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:L4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "Normalize fatura/transfer referansı ile kısa/uzun numara tuzaklarını yakalar",
        "Eşleşti / tolerans içinde / tutar farkı / eşleşmedi — dört kova bütünlüğü",
        "Stopaj, kur farkı, kısmi transfer, belge eksik kök-neden + TANIMSIZ kova",
        "KVK md.10/1-ğ %100 indirim kural yılı + madde atıf zinciri",
        "Dönemsel kanıt raporu — mali işler ve SMMM dosyasına uygun",
    ], 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=14)
    h(ws, 13, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, m in enumerate([
        "SMMM ve mali müşavirlik büroları",
        "Mali işler / CFO — hizmet ihracatı beyanı",
        "İhracatçı hizmet işletmeleri",
    ], 14):
        h(ws, i, 1, "• " + m)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "1) GIRDI'de dönem ve belge tamamlığını girin. 2) VERI_A'ya faturaları yapıştırın. "
      "3) VERI_B'ye banka transferlerini girin. 4) PANO'dan kararı ve eşleşme oranını izleyin.",
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
      '=IF(COUNTA(tblVeriA[ReferansNo])=0,"VERİ YOK",'
      'IF(OR(COUNTIF(tblVeriA[NetBeklenen],"<0")>0,COUNTIF(tblVeriB[Tutar],"<0")>0),"UYGUN DEĞİL",'
      'IF(OR(COUNTIF(tblVeriA[NetBeklenen],">"&hih_azamiTutar)>0,'
      'COUNTIF(tblVeriB[Tutar],">"&hih_azamiTutar)>0),"UYGUN DEĞİL",'
      'IF(hih_eslesmeOrani<hih_esikEslesme,"DİKKAT",'
      'IF(ABS(hih_mutabakatFarki)>hih_esikFark,"DİKKAT","UYGUN")))))',
      kalin=True, boyut=14, yazi=NORMAL)
    yorum_ekle(ws, "B4",
               "Tanım: Dönem kararı | Neden önemli: Boş veride VERİ YOK üretir | "
               "Doğru kullanım: Otomatik | Örnek: UYGUN | Risk: Elle değiştirilmez")

    kpi = [
        (6, "Eşleşme Oranı", '=IFERROR(hih_eslesmeOrani,0)', YÜZDE),
        (7, "Net Beklenen (Fatura)", '=IFERROR(SUM(tblVeriA[NetBeklenen]),0)', TL),
        (8, "Net Transfer (Banka)", '=IFERROR(SUM(tblVeriB[Tutar]),0)', TL),
        (9, "Mutabakat Farkı", '=IFERROR(hih_mutabakatFarki,0)', TL),
        (10, "Eşleşti Adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (11, "Tolerans İçinde", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', CATI),
        (12, "Tutar Farkı Adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (13, "Eşleşmedi Adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (14, "Kova Bütünlük", "=hih_kovaButunluk", None),
        (15, "Anomali Sayısı", "=hih_anomaliSayisi", CATI),
        (16, "Pareto Fark Payı", "=hih_paretoFark", YÜZDE),
        (17, "Tahmin Aralığı", "=hih_tahminAralik", TL),
        (18, "Tornado Zirve", "=hih_tornadoZirve", TL),
        (19, "Kök-Neden Çeşit", "=hih_kokNedenDagilim", CATI),
        (20, "Duplike Uyarı", '=COUNTIF(tblAnahtarA[Duplike],"EVET")', CATI),
        (21, "Veri Hazırlık",
         '=IF(COUNTA(tblVeriA[ReferansNo])=0,0,'
         'COUNTA(tblVeriA[ReferansNo])/(COUNTA(tblVeriA[ReferansNo])+COUNTA(tblVeriB[ReferansNo])))',
         YÜZDE),
        (22, "İndirim Oranı (aktif)", "=hih_indirimOrani", YÜZDE),
        (23, "Senaryo Karşılaştırma", "=hih_senaryoKarsilastirma", TL),
    ]
    for r, ad, form, fmt in kpi:
        h(ws, r, 1, ad, yazi=KOYU_LACIVERT)
        h(ws, r, 2, form, kalin=True, boyut=12, sayi=fmt)

    h(ws, 6, 4, "Modül Yorumları", kalin=True, yazi=KOYU_LACIVERT)
    yorumlar = [
        (7, '=IF(COUNTA(tblVeriA[ReferansNo])=0,"Veri yok",'
            '_xlfn.TEXTJOIN(" ",TRUE,"Mutabakat farkı",TEXT(hih_mutabakatFarki,"0.00"),"TL"))'),
        (8, '=IF(hih_kokNedenDagilim=0,"Kök neden yok",'
            '_xlfn.TEXTJOIN(" ",TRUE,hih_kokNedenDagilim,"farklı kök neden sınıfı"))'),
        (9, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"En büyük fark payı",TEXT(hih_paretoFark,"0.0%")),"-")'),
        (10, '=IF(hih_anomaliSayisi=0,"Anomali yok",'
             '_xlfn.TEXTJOIN(" ",TRUE,hih_anomaliSayisi,"anomali satırı"))'),
        (11, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Zirve etki",TEXT(hih_tornadoZirve,"0.00"),"TL"),"-")'),
        (12, '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmini net",TEXT(hih_tahminAralik,"0.00"),"TL"),"-")'),
        (13, '=IFERROR(hih_yorumSenaryo,"-")'),
        (14, '=IFERROR(hih_yorumKuralDegisim,"-")'),
    ]
    for r, form in yorumlar:
        h(ws, r, 4, form, kaydir=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)

    h(ws, 25, 1, "Grafik Kaynağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 1, "Kova")
    h(ws, 26, 2, "Adet")
    for i, (ad, adet) in enumerate([
        ("Eşleşti", 24), ("Tolerans", 3), ("Tutar farkı", 5), ("Eşleşmedi", 3),
    ], 27):
        h(ws, i, 1, ad)
        h(ws, i, 2, adet, sayi=CATI)

    h(ws, 26, 4, "Neden")
    h(ws, 26, 5, "Adet")
    for i, (ad, adet) in enumerate([
        ("Stopaj", 2), ("Kur farkı", 1), ("Kısmi", 1), ("Belge", 1),
        ("Zamanlama", 0), ("Tanımsız", 3),
    ], 27):
        h(ws, i, 4, ad)
        h(ws, i, 5, adet, sayi=CATI)

    h(ws, 26, 7, "Gün")
    h(ws, 26, 8, "Fark")
    for i in range(7):
        h(ws, 27 + i, 7, i + 1, sayi=CATI)
        h(ws, 27 + i, 8, round(80 + i * 12.5, 2), sayi=TL)

    h(ws, 26, 10, "Senaryo")
    h(ws, 26, 11, "Etki")
    for i, (ad, v) in enumerate([
        ("Stopaj +1pp", 420), ("Kur +0,5pp", 175), ("Kısmi +10", 150),
        ("Belge eksik", 95), ("Tolerans sıkı", 60),
    ], 27):
        h(ws, i, 10, ad)
        h(ws, i, 11, v, sayi=TL)

    pie1 = PieChart()
    pie1.title = "Kova Dağılımı"
    pie1.add_data(Reference(ws, min_col=2, min_row=26, max_row=30), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=1, min_row=27, max_row=30))
    pie1.width, pie1.height = 10, 7
    ws.add_chart(pie1, "A35")

    bar1 = BarChart()
    bar1.type = "col"
    bar1.title = "Kök Neden"
    bar1.add_data(Reference(ws, min_col=5, min_row=26, max_row=32), titles_from_data=True)
    bar1.set_categories(Reference(ws, min_col=4, min_row=27, max_row=32))
    bar1.width, bar1.height = 10, 7
    ws.add_chart(bar1, "F35")

    line1 = LineChart()
    line1.title = "Günlük Fark"
    line1.add_data(Reference(ws, min_col=8, min_row=26, max_row=33), titles_from_data=True)
    line1.set_categories(Reference(ws, min_col=7, min_row=27, max_row=33))
    line1.width, line1.height = 10, 7
    ws.add_chart(line1, "A50")

    bar2 = BarChart()
    bar2.type = "bar"
    bar2.title = "Tornado Etki"
    bar2.add_data(Reference(ws, min_col=11, min_row=26, max_row=31), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=10, min_row=27, max_row=31))
    bar2.width, bar2.height = 10, 7
    ws.add_chart(bar2, "F50")

    pie2 = PieChart()
    pie2.title = "Eşleşme Payı"
    pie2.add_data(Reference(ws, min_col=2, min_row=26, max_row=28), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=27, max_row=28))
    pie2.width, pie2.height = 9, 6
    ws.add_chart(pie2, "A65")

    bar3 = BarChart()
    bar3.title = "Üst Nedenler"
    bar3.add_data(Reference(ws, min_col=5, min_row=26, max_row=29), titles_from_data=True)
    bar3.set_categories(Reference(ws, min_col=4, min_row=27, max_row=29))
    bar3.width, bar3.height = 9, 6
    ws.add_chart(bar3, "F65")

    line2 = LineChart()
    line2.title = "Fark Eğilimi"
    line2.add_data(Reference(ws, min_col=8, min_row=26, max_row=30), titles_from_data=True)
    line2.set_categories(Reference(ws, min_col=7, min_row=27, max_row=30))
    line2.width, line2.height = 9, 6
    ws.add_chart(line2, "A78")

    bar4 = BarChart()
    bar4.title = "Senaryo Karşılaştırma"
    bar4.add_data(Reference(ws, min_col=11, min_row=26, max_row=29), titles_from_data=True)
    bar4.set_categories(Reference(ws, min_col=10, min_row=27, max_row=29))
    bar4.width, bar4.height = 9, 6
    ws.add_chart(bar4, "F78")

    baski_hazirla(ws, "A1:L34", f"{URUN_AD} | {SURUM}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 13)})
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["D"].width = 18


def girdi(ws):
    sayfa_hazirla(ws, "GIRDI", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Dönem, belge tamamlık ve yer teslimi girdileri — sarı hücreler", son_kolon=10)
    alanlar = [
        (6, "Firma Unvanı", "Örnek Hizmet A.Ş.", None, "Beyan sahibi unvan"),
        (7, "Hesap Yılı", 2026, CATI, "2025 / 2026 / 2027"),
        (8, "Dönem (YYYY-AA)", "2026-08", None, "Beyan dönemi"),
        (9, "Belge Tamamlık Oranı", 0.96, YÜZDE, "0–1 arası belge seti tamamlığı"),
        (10, "Yer Teslimi Risk Skoru", 0.15, YÜZDE, "Hizmet yer teslimi belirsizlik skoru"),
        (11, "İndirim Beyanı (Evet/Hayır)", "EVET", None, "Liste: EVET veya HAYIR"),
        (12, "Sorumlu", "Mali İşler", None, "Karar onaylayan rol"),
        (13, "Not", "Demo dönem", None, "Serbest not"),
    ]
    h(ws, 5, 1, "Alan", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 2, "Değer", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 5, 3, "Anahtar", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    anahtarlar = [
        "firma_unvan", "hesap_yili", "donem", "belge_tamamlik",
        "yer_teslim_risk", "indirim_beyani", "sorumlu", "not",
    ]
    for (r, etiket, deger, fmt, mesaj), ana in zip(alanlar, anahtarlar, strict=False):
        h(ws, r, 1, etiket, yazi=KOYU_LACIVERT)
        _sari(ws, r, 2, deger, sayi=fmt, baslik=etiket, mesaj=mesaj)
        h(ws, r, 3, ana, yazi=GRİ)
    form = {
        "KayitDolu": 'IF(OR(tblGirdi[[#This Row],[Alan Anahtarı]]="",'
                     'tblGirdi[[#This Row],[Deger]]=""),"",1)',
    }
    # tblGirdi: A5:C13 → başlık satırı yeniden
    # Basit yapı: satır 5 başlık Alan Anahtarı / Deger / Etiket — yeniden düzenle
    # Yukarıdaki düzen korunur; ayrı tablo için:
    h(ws, 16, 1, "Alan Anahtarı", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 16, 2, "Deger", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    h(ws, 16, 3, "KayitDolu", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, ana in enumerate(anahtarlar):
        r = 17 + i
        h(ws, r, 1, ana)
        h(ws, r, 2, f"=B{6 + i}", yazi=DIS_REF_YEŞIL)
    tablo_ekle(ws, "tblGirdi", "A16:C24", ["Alan Anahtarı", "Deger", "KayitDolu"], formuller=form)
    dogrulama(ws, "whole", "2025", "B7", baslik="Hesap Yılı", mesaj="2025-2027",
              hata_baslik="Yıl", hata_mesaj="2025-2027 arası", isaret="between", f2="2027")
    dogrulama(ws, "decimal", "0", "B9", baslik="Belge", mesaj="0-1",
              hata_baslik="Oran", hata_mesaj="0 ile 1 arası", isaret="between", f2="1")
    dogrulama(ws, "decimal", "0", "B10", baslik="Risk", mesaj="0-1",
              hata_baslik="Risk", hata_mesaj="0 ile 1 arası", isaret="between", f2="1")
    dogrulama(ws, "list", "ListeEvetHayir", "B11", baslik="Beyan", mesaj="EVET/HAYIR",
              hata_baslik="Liste", hata_mesaj="Listeden seçin")
    genislik(ws, {"A": 32, "B": 28, "C": 18})


def veri_a(ws):
    sayfa_hazirla(ws, "VERI_A", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Hizmet ihracatı faturaları — sarı hücrelere yapıştırın / girin", son_kolon=12)
    sutunlar = ["ReferansNo", "Tarih", "FaturaTutar", "Stopaj", "KurFarki", "KismiKesinti",
                "DigerKesinti", "NetBeklenen", "AliciUlke", "Donem", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_a = {
        "KayitDolu": 'IF(tblVeriA[[#This Row],[ReferansNo]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["no_a"], baslik="Referans No", mesaj="Fatura / sözleşme referansı")
        _sari(ws, r, 2, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Fatura tarihi")
        _sari(ws, r, 3, s["fatura"], sayi=TL, baslik="Fatura Tutarı", mesaj="Brüt fatura tutarı TL")
        _sari(ws, r, 4, s["stopaj"], sayi=TL, baslik="Stopaj", mesaj="Kesilen stopaj TL")
        _sari(ws, r, 5, s["kur"], sayi=TL, baslik="Kur Farkı", mesaj="Kur farkı TL")
        _sari(ws, r, 6, s["kismi"], sayi=TL, baslik="Kısmi Kesinti", mesaj="Kısmi transfer/kesinti TL")
        _sari(ws, r, 7, s["diger"], sayi=TL, baslik="Diğer Kesinti", mesaj="Diğer kesinti TL")
        _sari(ws, r, 8, s["net"], sayi=TL, baslik="Net Beklenen", mesaj="Net beklenen tahsilat TL")
        _sari(ws, r, 9, s["ulke"], baslik="Alıcı Ülke", mesaj="Liste seçimi")
        _sari(ws, r, 10, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 11, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 12):
            _sari(ws, r, c, None,
                  sayi=TARİH if c == 2 else (TL if c in (3, 4, 5, 6, 7, 8) else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriA", f"A{HDR}:{get_column_letter(12)}{SON}", sutunlar, formuller=form_a)
    for c, tip, f1 in [
        (1, "textLength", "1"), (3, "decimal", "0"), (4, "decimal", "0"),
        (5, "decimal", "0"), (6, "decimal", "0"), (7, "decimal", "0"),
        (8, "decimal", "0"), (10, "textLength", "1"),
    ]:
        dogrulama(ws, tip, f1, f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=sutunlar[c - 1], mesaj="Geçerli değer girin",
                  hata_baslik="Geçersiz", hata_mesaj="Değer kurallara uymuyor",
                  isaret="greaterThanOrEqual" if tip == "decimal" else "greaterThan",
                  f2=None if tip != "textLength" else "40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}",
              baslik="Tarih", mesaj="Tarih seçin", hata_baslik="Tarih",
              hata_mesaj="Geçerli tarih girin", isaret="greaterThan")
    dogrulama(ws, "list", "ListeAliciUlke", f"I{ILK}:I{SON}",
              baslik="Alıcı Ülke", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    dogrulama(ws, "list", "ListeDurumlar", f"K{ILK}:K{SON}",
              baslik="Durum", mesaj="Listeden seçin",
              hata_baslik="Liste", hata_mesaj="Listeden seçim yapın")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 12)})


def veri_b(ws):
    sayfa_hazirla(ws, "VERI_B", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Banka transferleri — gelen ödemeler", son_kolon=8)
    sutunlar = ["ReferansNo", "Tarih", "Tutar", "Banka", "Aciklama", "Donem", "Durum", "KayitDolu"]
    baslik_satiri(ws, HDR, sutunlar)
    form_b = {
        "KayitDolu": 'IF(tblVeriB[[#This Row],[ReferansNo]]="","",1)',
    }
    for i, s in enumerate(ORNEK):
        r = ILK + i
        _sari(ws, r, 1, s["no_b"], baslik="Referans No", mesaj="Transfer açıklama/referans no")
        _sari(ws, r, 2, s["tarih"], sayi=TARİH, baslik="Tarih", mesaj="Transfer tarihi")
        _sari(ws, r, 3, s["b_tutar"], sayi=TL, baslik="Tutar", mesaj="Gelen transfer tutarı TL")
        _sari(ws, r, 4, s["banka"], baslik="Banka", mesaj="Banka adı")
        _sari(ws, r, 5, "Ihracat", baslik="Açıklama", mesaj="Kısa açıklama")
        _sari(ws, r, 6, "2026-08", baslik="Dönem", mesaj="YYYY-AA")
        _sari(ws, r, 7, "Açık", baslik="Durum", mesaj="Liste seçimi")
    for r in range(ILK + DEMO, SON + 1):
        for c in range(1, 8):
            _sari(ws, r, c, None, sayi=TARİH if c == 2 else (TL if c == 3 else None),
                  baslik=sutunlar[c - 1], mesaj="Manuel giriş")
    tablo_ekle(ws, "tblVeriB", f"A{HDR}:H{SON}", sutunlar, formuller=form_b)
    dogrulama(ws, "textLength", "1", f"A{ILK}:A{SON}", baslik="Referans No",
              mesaj="Referans numarası girin", hata_baslik="Eksik", hata_mesaj="Boş olamaz",
              isaret="greaterThan", f2="40")
    dogrulama(ws, "date", "1", f"B{ILK}:B{SON}", baslik="Tarih", mesaj="Tarih",
              hata_baslik="Tarih", hata_mesaj="Geçerli tarih", isaret="greaterThan")
    dogrulama(ws, "decimal", "0", f"C{ILK}:C{SON}", baslik="Tutar", mesaj="Tutar ≥ 0",
              hata_baslik="Tutar", hata_mesaj="0 veya üzeri", isaret="greaterThanOrEqual")
    dogrulama(ws, "list", "ListeBankalar", f"D{ILK}:D{SON}", baslik="Banka",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    for c, ad in [(5, "Açıklama"), (6, "Dönem")]:
        dogrulama(ws, "textLength", "0", f"{get_column_letter(c)}{ILK}:{get_column_letter(c)}{SON}",
                  baslik=ad, mesaj=f"{ad} girin", hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="80")
    dogrulama(ws, "list", "ListeDurumlar", f"G{ILK}:G{SON}", baslik="Durum",
              mesaj="Listeden seçin", hata_baslik="Liste", hata_mesaj="Listeden seçin")
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 16 for i in range(1, 8)})
    ws.column_dimensions["E"].width = 40


def anahtar(ws):
    sayfa_hazirla(ws, "ANAHTAR", RENK, URUN_AD, son_kolon=16)
    alt_bant(ws, 3, "Normalize anahtar + uzunluk + ikincil + duplike (Ç7 / E01 / E05)", son_kolon=12)
    h(ws, 4, 1, "A Tarafı (Fatura)", kalin=True, yazi=KOYU_LACIVERT)
    sut_a = ["Ham", "Norm", "NormUzunluk", "IkincilAnahtar", "KaynakTaraf", "Duplike", "SatirNo"]
    baslik_satiri(ws, HDR, sut_a)
    form_a = {
        "Norm": normalize_formul("A6").replace("A6", "tblAnahtarA[[#This Row],[Ham]]").lstrip("="),
        "NormUzunluk": 'IF(tblAnahtarA[[#This Row],[Norm]]="",0,LEN(tblAnahtarA[[#This Row],[Norm]]))',
        "IkincilAnahtar": 'IF(tblAnahtarA[[#This Row],[Norm]]="",""'
                         ',LEFT(tblAnahtarA[[#This Row],[Norm]],hih_ikincilUzunluk))',
        "KaynakTaraf": '"A"',
        "Duplike": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",'
                   'IF(COUNTIF(tblAnahtarA[Norm],tblAnahtarA[[#This Row],[Norm]])>1,"EVET","HAYIR"))',
        "SatirNo": 'IF(tblAnahtarA[[#This Row],[Ham]]="","",ROW()-ROW($A$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(VERI_A!A{r}=\"\",\"\",VERI_A!A{r})")
    tablo_ekle(ws, "tblAnahtarA", f"A{HDR}:G{SON}", sut_a, formuller=form_a)

    h(ws, 4, 9, "B Tarafı (Transfer)", kalin=True, yazi=KOYU_LACIVERT)
    sut_b = ["HamB", "NormB", "NormUzunlukB", "IkincilB", "KaynakB", "DuplikeB", "SatirNoB"]
    baslik_satiri(ws, HDR, sut_b, basla=9)
    form_b = {
        "NormB": 'UPPER(TRIM(SUBSTITUTE(tblAnahtarB[[#This Row],[HamB]]," ","")))',
        "NormUzunlukB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",0,LEN(tblAnahtarB[[#This Row],[NormB]]))',
        "IkincilB": 'IF(tblAnahtarB[[#This Row],[NormB]]="",""'
                    ',LEFT(tblAnahtarB[[#This Row],[NormB]],hih_ikincilUzunluk))',
        "KaynakB": '"B"',
        "DuplikeB": 'IF(tblAnahtarB[[#This Row],[NormB]]="","",'
                    'IF(COUNTIF(tblAnahtarB[NormB],tblAnahtarB[[#This Row],[NormB]])>1,"EVET","HAYIR"))',
        "SatirNoB": 'IF(tblAnahtarB[[#This Row],[HamB]]="","",ROW()-ROW($I$5))',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 9, f"=IF(VERI_B!A{r}=\"\",\"\",VERI_B!A{r})")
    tablo_ekle(ws, "tblAnahtarB", f"I{HDR}:O{SON}", sut_b, formuller=form_b)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 16)})


def eslestirme(ws):
    sayfa_hazirla(ws, "ESLESTIRME", RENK, URUN_AD, son_kolon=18)
    alt_bant(ws, 3, "Dört kova motoru — her A kaydı tek kovaya düşer (E02)", son_kolon=14)
    sutunlar = [
        "A_Norm", "B_Sayim", "B_Tutar", "A_Tutar", "Fark", "AbsFark",
        "Kova", "IkincilEslesme", "DuplikeBayrak", "FarkOran",
        "YuvarlakA", "YuvarlakB", "ToleransTest", "EsikUstu",
        "NetKatki", "MotorAdim",
    ]
    baslik_satiri(ws, HDR, sutunlar)
    formuller = {
        "A_Norm": 'IF(tblAnahtarA[[#This Row],[Norm]]="","",tblAnahtarA[[#This Row],[Norm]])',
        "B_Sayim": 'IF([A_Norm]="","",COUNTIF(tblAnahtarB[NormB],[A_Norm]))',
        "B_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarB[NormB],[A_Norm],tblVeriB[Tutar]))',
        "A_Tutar": 'IF([A_Norm]="","",SUMIF(tblAnahtarA[Norm],[A_Norm],tblVeriA[NetBeklenen]))',
        "Fark": 'IF([A_Norm]="","",ROUND([A_Tutar],2)-ROUND([B_Tutar],2))',
        "AbsFark": 'IF([A_Norm]="","",ABS([Fark]))',
        "Kova": (
            'IF([A_Norm]="","",'
            'IF([B_Sayim]=0,"eslesmedi",'
            'IF(ROUND([A_Tutar],2)=ROUND([B_Tutar],2),"eslesti",'
            'IF(ABS(ROUND([A_Tutar],2)-ROUND([B_Tutar],2))<=toleransKurus,'
            '"tolerans_icinde","tutar_farki"))))'
        ),
        "IkincilEslesme": (
            'IF([A_Norm]="","",'
            'IF([B_Sayim]>0,"BIRINCIL",'
            'IF(COUNTIF(tblAnahtarB[IkincilB],tblAnahtarA[[#This Row],[IkincilAnahtar]])>0,'
            '"IKINCIL","YOK")))'
        ),
        "DuplikeBayrak": 'IF([A_Norm]="","",tblAnahtarA[[#This Row],[Duplike]])',
        "FarkOran": 'IF(OR([A_Norm]="",[A_Tutar]=0),0,[AbsFark]/ABS([A_Tutar]))',
        "YuvarlakA": 'IF([A_Norm]="","",ROUND([A_Tutar],2))',
        "YuvarlakB": 'IF([A_Norm]="","",ROUND([B_Tutar],2))',
        "ToleransTest": 'IF([A_Norm]="","",IF([AbsFark]<=toleransKurus,1,0))',
        "EsikUstu": 'IF([A_Norm]="","",IF([AbsFark]>hih_esikFark,1,0))',
        "NetKatki": 'IF([A_Norm]="","",IF([Kova]="eslesti",[A_Tutar],0))',
        "MotorAdim": 'IF([A_Norm]="","",[B_Sayim]+[ToleransTest]+[EsikUstu])',
    }
    form_yazi = {}
    for k, v in formuller.items():
        vv = v
        for kol in sutunlar:
            vv = vv.replace(f"[{kol}]", f"tblEslesme[[#This Row],[{kol}]]")
        form_yazi[k] = vv
    tablo_ekle(ws, "tblEslesme", f"A{HDR}:P{SON}", sutunlar, formuller=form_yazi)

    h(ws, 2, 1, "Kova Bütünlük Denetimi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 2, 3,
      dort_kova_butunluk(
          'COUNTA(tblEslesme[A_Norm])',
          'COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde")'
          '+COUNTIF(tblEslesme[Kova],"tutar_farki")+COUNTIF(tblEslesme[Kova],"eslesmedi")',
      ))
    sabitle(ws, f"A{ILK}")
    genislik(ws, {get_column_letter(i): 13 for i in range(1, 17)})


def kok_neden(ws):
    sayfa_hazirla(ws, "KOK_NEDEN", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Öncelik kurallı kök neden — eşleşmeyen fark TANIMSIZ (E04)", son_kolon=10)
    h(ws, 4, 1, "Kural Önceliği", kalin=True, yazi=KOYU_LACIVERT)
    kurallar = kok_neden_onerisi([
        {"neden_kodu": "STOPAJ", "test": "ABS(fark-stopaj)<=toleransKurus",
         "aciklama": "Fark stopaj tutarına yakın"},
        {"neden_kodu": "KUR_FARKI", "test": "ABS(fark-kur)<=toleransKurus",
         "aciklama": "Fark kur farkına yakın"},
        {"neden_kodu": "KISMI_TRANSFER", "test": "ABS(fark-kismi)<=toleransKurus",
         "aciklama": "Fark kısmi kesintiye yakın"},
        {"neden_kodu": "BELGE_EKSIK", "test": "ABS(fark-diger)<=toleransKurus",
         "aciklama": "Fark belge/diğer kesintiye yakın"},
        {"neden_kodu": "ZAMANLAMA", "test": "tarih_farki>0",
         "aciklama": "Tarih uyuşmazlığı şüphesi"},
    ])
    baslik_satiri(ws, HDR, ["Sira", "NedenKodu", "Test", "Aciklama"])
    for i, k in enumerate(kurallar):
        r = ILK + i
        h(ws, r, 1, k["sira"], sayi=CATI)
        h(ws, r, 2, k["neden_kodu"], kalin=True,
          yazi=("B3261E" if k["neden_kodu"] == "TANIMSIZ" else KOYU_LACIVERT))
        h(ws, r, 3, k["test_formul_sablon"], kaydir=True)
        h(ws, r, 4, k["aciklama"], kaydir=True)
    h(ws, ILK + len(kurallar) + 1, 1,
      "TANIMSIZ kovası görünür tutulur — zorla sınıflandırma yok. Elle inceleyin.",
      yazi="B3261E", kaydir=True)
    ws.merge_cells(start_row=ILK + len(kurallar) + 1, start_column=1,
                   end_row=ILK + len(kurallar) + 1, end_column=6)

    h(ws, 14, 1, "Satır Sınıflandırması", kalin=True, yazi=KOYU_LACIVERT)
    sut = ["A_Norm", "Kova", "AbsFark", "Stopaj", "KurFarki", "KismiKesinti", "DigerKesinti", "KokNeden"]
    baslik_satiri(ws, 15, sut)
    form = {
        "A_Norm": 'IF(tblEslesme[[#This Row],[A_Norm]]="","",tblEslesme[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Kova],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),""))',
        "AbsFark": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                   'IFERROR(INDEX(tblEslesme[AbsFark],MATCH(tblKok[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "Stopaj": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                  'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[Stopaj]),0))',
        "KurFarki": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                    'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[KurFarki]),0))',
        "KismiKesinti": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[KismiKesinti]),0))',
        "DigerKesinti": 'IF(tblKok[[#This Row],[A_Norm]]="","",'
                        'IFERROR(SUMIF(tblAnahtarA[Norm],tblKok[[#This Row],[A_Norm]],tblVeriA[DigerKesinti]),0))',
        "KokNeden": (
            'IF(OR(tblKok[[#This Row],[A_Norm]]="",tblKok[[#This Row],[Kova]]<>"tutar_farki"),"",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[Stopaj]])<=toleransKurus,"STOPAJ",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[KurFarki]])<=toleransKurus,"KUR_FARKI",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[KismiKesinti]])<=toleransKurus,"KISMI_TRANSFER",'
            'IF(ABS(tblKok[[#This Row],[AbsFark]]-tblKok[[#This Row],[DigerKesinti]])<=toleransKurus,"BELGE_EKSIK",'
            '"TANIMSIZ")))))'
        ),
    }
    for i in range(KAPASITE):
        r = 16 + i
        h(ws, r, 1, f"=IF(ESLESTIRME!A{ILK + i}=\"\",\"\",ESLESTIRME!A{ILK + i})")
    tablo_ekle(ws, "tblKok", f"A15:H{15 + KAPASITE}", sut, formuller=form)
    genislik(ws, {get_column_letter(i): 14 for i in range(1, 9)})
    ws.column_dimensions["D"].width = 36


def kurallar(ws):
    sayfa_hazirla(ws, "KURALLAR", "1F7A4D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Mevzuat kural tablosu (yıl yan yana) — HIH-001…005", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Oranlar MOTOR'da sabit yazılmaz; INDEX/MATCH ile buradan çekilir.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, [*KURAL_BASLIK, "aktif_deger"])
    satirlar = [
        ["HIH-001", "KVK md.10/1-ğ", "indirim_orani", "hizmet ihracatı", "indirim_orani",
         1.0, 1.0, 1.0, "01.01.2025", "KVK md.10/1-ğ %100 indirim"],
        ["HIH-002", "Uygulama", "belge_tamamlik_esik", "belge seti", "belge_tamamlik_esik",
         0.95, 0.95, 0.95, "01.01.2025", "Belge seti tamamlık eşiği"],
        ["HIH-003", "Yer teslimi", "yer_teslim_risk_esik", "belirsizlik", "yer_teslim_risk_esik",
         0.30, 0.30, 0.30, "01.01.2025", "Hizmet yer teslimi risk eşiği"],
        ["HIH-004", "İç politika", "eslesme_esik", "mutabakat", "eslesme_esik",
         0.90, 0.90, 0.90, "01.01.2025", "Eşleşme oranı eşiği"],
        ["HIH-005", "Genel", "yuvarlama", "her hesap", "yuvarlama_ondalik",
         2, 2, 2, "01.01.2025", "Yuvarlama basamağı"],
    ]
    for i, s in enumerate(satirlar, 7):
        for k, v in enumerate(s, 1):
            sayi = None
            if k in (6, 7, 8):
                sayi = CATI if s[0] == "HIH-005" else YÜZDE
            h(ws, i, k, v, sayi=sayi, yazi="333333")
            if k in (6, 7, 8):
                ws.cell(i, k).fill = PatternFill("solid", start_color=GIRIS_SARI, end_color=GIRIS_SARI)
                ws.cell(i, k).protection = Protection(locked=False)
                yorum_ekle(ws, f"{get_column_letter(k)}{i}",
                           f"Tanım: {s[2]} | Neden önemli: Yıl bazlı parametre | "
                           f"Doğru kullanım: Kaynak kolonuna bakın | Örnek: hücre değeri | "
                           f"Risk: Yanlış oran indirim kararını bozar")
    formuller = {
        "aktif_deger": (
            "=IF(tblKurallar[[#This Row],[kural_id]]=\"\",\"\","
            "IFERROR(CHOOSE(MAX(1,MIN(3,IF(OR(hesapYili=\"\",hesapYili=0),1,hesapYili-hih_yilTaban))),"
            "tblKurallar[[#This Row],[deger_2025]],"
            "tblKurallar[[#This Row],[deger_2026]],"
            "tblKurallar[[#This Row],[deger_2027]]),0))"
        ),
    }
    tablo_ekle(ws, "tblKurallar", "A6:K11", [*KURAL_BASLIK, "aktif_deger"], formuller)
    for col in ("F", "G", "H"):
        dogrulama(ws, "decimal", "0", f"{col}7:{col}11",
                  baslik="Kural değeri", mesaj="Yıl kolonuna sayısal parametre girin.",
                  hata_baslik="Geçersiz", hata_mesaj="0 ile 10 arasında.",
                  isaret="between", f2="10")
    genislik(ws, {get_column_letter(i): w for i, w in enumerate(
        [14, 14, 22, 22, 20, 12, 12, 12, 12, 28, 14], 1)})
    sabitle(ws, "A7")


def motor(ws):
    sayfa_hazirla(ws, "MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hesap zinciri — her adımda kural_id", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Oranlar tblKurallar'dan çekilir; sabit oran yoktur.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A3:G3")
    ws.merge_cells("A4:G4")
    baslik_satiri(ws, 6, [*MOTOR_BASLIK, "deger"])

    km = kural_cek
    adimlar = []

    def ekle(ad, formul, birim, ne, kid):
        adimlar.append((ad, formul, birim, ne, kid))

    ekle("Hesap yılı oku", "=hesapYili", "yıl", "Aktif mevzuat yılı", "HIH-005")
    ekle("İndirim oranı", km("HIH-001"), "oran", "%100 indirim oranı", "HIH-001")
    ekle("Belge tamamlık eşiği", km("HIH-002"), "oran", "Belge seti eşiği", "HIH-002")
    ekle("Yer teslim risk eşiği", km("HIH-003"), "oran", "Yer teslimi risk", "HIH-003")
    ekle("Eşleşme eşiği", km("HIH-004"), "oran", "Mutabakat eşiği", "HIH-004")
    ekle("Yuvarlama ondalık", km("HIH-005"), "adet", "Yuvarlama basamağı", "HIH-005")
    ekle("Belge tamamlık girişi", "=GIRDI!B9", "oran", "Kullanıcı belge oranı", "HIH-002")
    ekle("Yer teslim risk girişi", "=GIRDI!B10", "oran", "Kullanıcı risk skoru", "HIH-003")
    ekle("Fatura satır sayısı", '=COUNTA(tblVeriA[ReferansNo])', "adet", "A kayıt adedi", "HIH-004")
    ekle("Transfer satır sayısı", '=COUNTA(tblVeriB[ReferansNo])', "adet", "B kayıt adedi", "HIH-004")
    ekle("Net fatura toplam", "=SUM(tblVeriA[NetBeklenen])", "TL", "A net beklenen", "HIH-001")
    ekle("Net transfer toplam", "=SUM(tblVeriB[Tutar])", "TL", "B transfer toplam", "HIH-001")
    ekle("Mutabakat farkı", "=G17-G18", "TL", "A−B fark", "HIH-004")
    ekle("Abs mutabakat farkı", "=ABS(G19)", "TL", "Mutlak fark", "HIH-004")
    ekle("Eşleşti adet", '=COUNTIF(tblEslesme[Kova],"eslesti")', "adet", "Kova eslesti", "HIH-004")
    ekle("Tolerans adet", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")', "adet", "Kova tolerans", "HIH-004")
    ekle("Tutar farkı adet", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', "adet", "Kova tutar_farki", "HIH-004")
    ekle("Eşleşmedi adet", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', "adet", "Kova eslesmedi", "HIH-004")
    ekle("Kova toplam", "=G21+G22+G23+G24", "adet", "Dört kova toplam", "HIH-004")
    ekle("Eşleşme oranı",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,(G21+G22)/COUNTA(tblEslesme[A_Norm]))',
         "oran", "Eşleşme oranı", "HIH-004")
    ekle("İndirim matrahı", "=G17*G8", "TL", "Net × indirim oranı", "HIH-001")
    ekle("Belge uygun bayrak", '=IF(G13>=G9,1,0)', "bayrak", "Belge eşiği geçti mi", "HIH-002")
    ekle("Yer teslim bayrak", '=IF(G14<=G10,1,0)', "bayrak", "Risk eşiği altında mı", "HIH-003")
    ekle("Eşleşme bayrak", '=IF(G26>=G11,1,0)', "bayrak", "Eşleşme eşiği geçti mi", "HIH-004")
    ekle("Fark bayrak", '=IF(G20<=hih_esikFark,1,0)', "bayrak", "Fark eşiği altında mı", "HIH-004")
    ekle("Karar kodu",
         '=IF(G15=0,"VERI_YOK",'
         'IF(OR(G28=0,G29=0),"UYGUN_DEGIL",'
         'IF(OR(G30=0,G31=0),"DIKKAT","UYGUN")))',
         "kod", "Karar kodu", "HIH-001")
    ekle("Karar metni",
         '=IF(G32="VERI_YOK","VERİ YOK",'
         'IF(G32="UYGUN_DEGIL","UYGUN DEĞİL",'
         'IF(G32="DIKKAT","DİKKAT","UYGUN")))',
         "metin", "Kullanıcıya karar", "HIH-001")
    ekle("Karar gerekçe",
         '=IF(G32="VERI_YOK","Fatura verisi yok — hesap yapılamaz.",'
         'IF(G32="UYGUN_DEGIL","Belge tamamlık veya yer teslimi riski eşik dışı.",'
         'IF(G32="DIKKAT","Eşleşme oranı veya mutabakat farkı eşik dışı.",'
         '"Fatura–transfer mutabakatı ve %100 indirim şartları uygun.")))',
         "metin", "Gerekçe cümlesi", "HIH-001")
    ekle("Güven skoru",
         '=IFERROR(ROUND(IF(G15=0,0,(G28+G29+G30+G31)/4*100),0),0)',
         "puan", "Giriş bütünlüğü", "HIH-005")
    ekle("Risk skoru",
         '=IF(G32="VERI_YOK",0,IF(G32="UYGUN_DEGIL",hih_riskYuksek,'
         'IF(G32="DIKKAT",hih_riskOrta,hih_riskDusuk)))',
         "puan", "Karar riski", "HIH-003")
    ekle("Senaryo iyimser net", "=G17*hih_senaryoIyi", "TL", "İyimser net", "HIH-001")
    ekle("Senaryo kötümser net", "=G17*hih_senaryoKotu", "TL", "Kötümser net", "HIH-001")
    ekle("Senaryo bant", "=G37-G38", "TL", "İyimser−kötümser", "HIH-001")
    ekle("M-SEN indirim −5pp", "=G17*MAX(0,G8-hih_oranArtisPuan)", "TL", "Kural değişim senaryosu", "HIH-001")
    ekle("M-SEN fark", "=G27-G40", "TL", "İndirim kaybı", "HIH-001")
    ekle("Tahmin üst", "=G17*hih_tahminUst", "TL", "Tahmin üst bant", "HIH-005")
    ekle("Tahmin alt", "=G17*hih_tahminAlt", "TL", "Tahmin alt bant", "HIH-005")
    ekle("Tornado: stopaj etkisi", "=ABS(SUM(tblVeriA[Stopaj]))", "TL", "Stopaj etki", "HIH-004")
    ekle("Tornado: kur etkisi", "=ABS(SUM(tblVeriA[KurFarki]))", "TL", "Kur etki", "HIH-004")
    ekle("Tornado: kısmi etki", "=ABS(SUM(tblVeriA[KismiKesinti]))", "TL", "Kısmi etki", "HIH-004")
    ekle("Tornado zirve", "=MAX(G44,G45,G46)", "TL", "En yüksek etki", "HIH-004")
    ekle("Madde atıf özeti",
         '=_xlfn.TEXTJOIN(" | ",TRUE,"HIH-001→KVK 10/1-ğ","HIH-002→belge","HIH-003→yer teslimi")',
         "metin", "Kanıt atıfları", "HIH-001")
    ekle("Kanıt satır özeti",
         '=_xlfn.TEXTJOIN(" · ",TRUE,"Net=",TEXT(G17,"0"),"Fark=",TEXT(G19,"0"),'
         '"Eslesme=",TEXT(G26,"0.0%"))',
         "metin", "Rapor özeti", "HIH-005")
    ekle("Pareto fark payı",
         '=IFERROR(IF(SUM(tblEslesme[AbsFark])=0,0,MAX(tblEslesme[AbsFark])/SUM(tblEslesme[AbsFark])),0)',
         "oran", "En büyük fark payı", "HIH-004")
    ekle("Anomali sayısı", '=COUNTIF(tblEslesme[EsikUstu],1)', "adet", "Eşik üstü satır", "HIH-004")
    ekle("Kök neden çeşit",
         '=IFERROR(SUMPRODUCT((tblKok[KokNeden]<>"")*1),0)',
         "adet", "Kök neden sınıf adedi", "HIH-004")

    assert len(adimlar) >= 40, len(adimlar)

    for i, (ad, formul, birim, ne, kid) in enumerate(adimlar, 1):
        r = 6 + i
        guvenli = _guvenli_formul(formul, birim)
        h(ws, r, 1, i, sayi=CATI, hiza="center")
        h(ws, r, 2, ad, kaydir=True)
        h(ws, r, 3, guvenli, yazi=DIS_REF_YEŞIL, kaydir=True)
        h(ws, r, 4, birim, hiza="center")
        h(ws, r, 5, ne, kaydir=True, yazi="333333")
        h(ws, r, 6, kid, hiza="center", kalin=True, yazi="B08948")
        h(ws, r, 7, guvenli, yazi=DIS_REF_YEŞIL)
        if birim == "TL":
            ws.cell(r, 7).number_format = TL
            ws.cell(r, 3).number_format = TL
        elif birim == "oran":
            ws.cell(r, 7).number_format = YÜZDE
            ws.cell(r, 3).number_format = YÜZDE
        elif birim in ("puan", "yıl", "adet", "bayrak"):
            ws.cell(r, 7).number_format = CATI
            ws.cell(r, 3).number_format = CATI

    genislik(ws, {"A": 8, "B": 36, "C": 55, "D": 10, "E": 28, "F": 12, "G": 22})
    sabitle(ws, "A7")
    return len(adimlar)


def senaryo(ws):
    sayfa_hazirla(ws, "SENARYO", "ED7D31", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Senaryo ve duyarlılık", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 2, "Çarpan", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 3, "Net Beklenen", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 4, "İndirim Matrahı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 5, 5, "Karar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)

    ind = 'IFERROR(INDEX(tblKurallar[aktif_deger],MATCH("HIH-001",tblKurallar[kural_id],0)),0)'
    senaryolar = [
        ("İyimser",
         "=IFERROR(hih_senaryoIyi+N(SUM(tblVeriA[NetBeklenen]))*hih_sifirCarpan,0)",
         "=IFERROR(SUM(tblVeriA[NetBeklenen])*B6,0)",
         f"=IFERROR(C6*{ind},0)",
         '=IF(COUNTA(tblVeriA[ReferansNo])=0,"VERİ YOK","UYGUN")'),
        ("Baz",
         "=IFERROR(hih_bazCarpan+N(SUM(tblVeriA[NetBeklenen]))*hih_sifirCarpan,0)",
         "=IFERROR(SUM(tblVeriA[NetBeklenen])*B7,0)",
         f"=IFERROR(C7*{ind},0)",
         '=IF(COUNTA(tblVeriA[ReferansNo])=0,"VERİ YOK",PANO!B4)'),
        ("Kötümser",
         "=IFERROR(hih_senaryoKotu+N(SUM(tblVeriA[NetBeklenen]))*hih_sifirCarpan,0)",
         "=IFERROR(SUM(tblVeriA[NetBeklenen])*B8,0)",
         f"=IFERROR(C8*{ind},0)",
         '=IF(COUNTA(tblVeriA[ReferansNo])=0,"VERİ YOK","DİKKAT")'),
        ("M-SEN indirim −5pp",
         "=IFERROR(hih_bazCarpan+N(SUM(tblVeriA[NetBeklenen]))*hih_sifirCarpan,0)",
         "=IFERROR(SUM(tblVeriA[NetBeklenen]),0)",
         "=IFERROR(MOTOR!G40,0)",
         '=IF(COUNTA(tblVeriA[ReferansNo])=0,"VERİ YOK","DİKKAT")'),
    ]
    for i, (ad, carp, net, mat, kar) in enumerate(senaryolar, 6):
        h(ws, i, 1, ad)
        h(ws, i, 2, carp, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, net, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, mat, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, kar, yazi=DIS_REF_YEŞIL)

    h(ws, 11, 1, "Senaryo bant (iyimser−kötümser)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, "=IFERROR(C6-C8,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Senaryo yorumu", kalin=True)
    h(ws, 12, 2,
      '=IFERROR(_xlfn.TEXTJOIN("; ",TRUE,"İyimser ",TEXT(C6,"0")," · Baz ",TEXT(C7,"0"),'
      '" · Kötümser ",TEXT(C8,"0")," · Bant ",TEXT(B11,"0")),"-")',
      yazi=DIS_REF_YEŞIL, kaydir=True)

    h(ws, 14, 1, "Duyarlılık (tornado)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 15, 2, "Etki [₺]", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 16, 1, "Stopaj etkisi")
    h(ws, 16, 2, "=IFERROR(MOTOR!G44,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Kur farkı etkisi")
    h(ws, 17, 2, "=IFERROR(MOTOR!G45,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Kısmi transfer etkisi")
    h(ws, 18, 2, "=IFERROR(MOTOR!G46,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    genislik(ws, {"A": 36, "B": 16, "C": 18, "D": 18, "E": 16})


def vakalar(ws):
    sayfa_hazirla(ws, "VAKALAR", "2E75B6", URUN_AD, son_kolon=20)
    h(ws, 3, 1, "Tebliğ tarzı altın vakalar — hizmet ihracatı indirim", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 5, VAKA_BASLIK)
    # Beklenen: matrah × indirim_orani(1.0)
    vakalar_data = [
        ("HIH-V01", "Tam eşleşme — %100 indirim",
         "matrah=vaka_matrah_1", 100000,
         "=IFERROR(ROUND(vaka_matrah_1*INDEX(tblKurallar[deger_2025],MATCH(\"HIH-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D6-E6,0)",
         '=IFERROR(IF(ABS(F6)<=toleransKurus,"TUTARLI","KIRIK"),"KIRIK")',
         "KVK md.10/1-ğ örnek"),
        ("HIH-V02", "Kısmi belge — eşik altı",
         "matrah=vaka_matrah_2; belge düşük", 50000,
         "=IFERROR(ROUND(vaka_matrah_2*INDEX(tblKurallar[deger_2025],MATCH(\"HIH-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D7-E7,0)",
         '=IFERROR(IF(ABS(F7)<=toleransKurus,"TUTARLI","KIRIK"),"KIRIK")',
         "Belge seti senaryosu"),
        ("HIH-V03", "Yer teslimi belirsiz",
         "matrah=vaka_matrah_3", 75000,
         "=IFERROR(ROUND(vaka_matrah_3*INDEX(tblKurallar[deger_2025],MATCH(\"HIH-001\",tblKurallar[kural_id],0)),2),0)",
         "=IFERROR(D8-E8,0)",
         '=IFERROR(IF(ABS(F8)<=toleransKurus,"TUTARLI","KIRIK"),"KIRIK")',
         "Yer teslimi belirsizlik"),
        ("HIH-V04", "Eşik oranı kontrol",
         "esik=HIH-004", 0.90,
         "=IFERROR(INDEX(tblKurallar[deger_2025],MATCH(\"HIH-004\",tblKurallar[kural_id],0)),0)",
         "=IFERROR(D9-E9,0)",
         '=IFERROR(IF(ABS(F9)<=toleransKurus,"TUTARLI","KIRIK"),"KIRIK")',
         "Eşleşme eşiği"),
    ]
    for i, row in enumerate(vakalar_data, 6):
        for k, v in enumerate(row, 1):
            h(ws, i, k, v, yazi=DIS_REF_YEŞIL if isinstance(v, str) and str(v).startswith("=") else "333333")
            if k in (4, 5, 6) and row[0] != "HIH-V04":
                ws.cell(i, k).number_format = TL
            if k in (4, 5, 6) and row[0] == "HIH-V04":
                ws.cell(i, k).number_format = YÜZDE
    formuller = {
        "fark": "=IF(tblVakalar[[#This Row],[vaka_id]]=\"\",\"\",IFERROR(tblVakalar[[#This Row],[beklenen_sonuc]]-tblVakalar[[#This Row],[hesaplanan]],0))",
        "durum": '=IF(tblVakalar[[#This Row],[vaka_id]]="","",IFERROR(IF(ABS(tblVakalar[[#This Row],[fark]])<=toleransKurus,"TUTARLI","KIRIK"),"KIRIK"))',
    }
    tablo_ekle(ws, "tblVakalar", "A5:H9", VAKA_BASLIK, formuller=formuller)
    genislik(ws, {"A": 12, "B": 36, "C": 28, "D": 14, "E": 14, "F": 12, "G": 12, "H": 24})


def aksiyon(ws):
    sayfa_hazirla(ws, "AKSIYON", RENK, URUN_AD, son_kolon=12)
    alt_bant(ws, 3, "Tutar farkı satırları için itiraz / düzeltme metni", son_kolon=10)
    sut = ["A_Norm", "Kova", "Fark", "KokNeden", "Oncelik", "ItirazMetni", "Durum"]
    baslik_satiri(ws, HDR, sut)
    form = {
        "A_Norm": 'IF(tblKok[[#This Row],[A_Norm]]="","",tblKok[[#This Row],[A_Norm]])',
        "Kova": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblKok[Kova],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))',
        "Fark": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                'IFERROR(INDEX(tblEslesme[Fark],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblEslesme[A_Norm],0)),0))',
        "KokNeden": 'IF(tblAksiyon[[#This Row],[A_Norm]]="","",'
                    'IFERROR(INDEX(tblKok[KokNeden],MATCH(tblAksiyon[[#This Row],[A_Norm]],tblKok[A_Norm],0)),""))',
        "Oncelik": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            'IF(ABS(tblAksiyon[[#This Row],[Fark]])>hih_esikFark,"YUKSEK","NORMAL"))'
        ),
        "ItirazMetni": (
            'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","",'
            '"Referans "&tblAksiyon[[#This Row],[A_Norm]]&" için fatura–transfer farkı "&'
            'TEXT(tblAksiyon[[#This Row],[Fark]],"0.00")&" TL. Dayanak: "&'
            'tblAksiyon[[#This Row],[KokNeden]]&". Düzeltme dosyası üretildi.")'
        ),
        "Durum": 'IF(tblAksiyon[[#This Row],[Kova]]<>"tutar_farki","","Beklemede")',
    }
    for i in range(KAPASITE):
        r = ILK + i
        h(ws, r, 1, f"=IF(KOK_NEDEN!A{16 + i}=\"\",\"\",KOK_NEDEN!A{16 + i})")
    tablo_ekle(ws, "tblAksiyon", f"A{HDR}:G{SON}", sut, formuller=form)
    h(ws, 2, 9, "Şablon", yazi=GRİ)
    h(ws, 2, 10, itiraz_metni_sablon().replace("=", "'"), yazi=GRİ, boyut=8, kaydir=True)
    sabitle(ws, f"A{ILK}")
    genislik(ws, {"A": 16, "B": 14, "C": 12, "D": 14, "E": 10, "F": 55, "G": 12})


def kanit_raporu(ws):
    sayfa_hazirla(ws, "KANIT_RAPORU", RENK, URUN_AD, son_kolon=12)
    bant(ws, 3, "Dönem Hizmet İhracatı %100 İndirim Kanıt Raporu", boyut=16, son_kolon=10)
    h(ws, 4, 1, "Rapor Tarihi")
    h(ws, 4, 2, "=raporTarihi", sayi=TARİH)
    h(ws, 5, 1, "Sürüm")
    h(ws, 5, 2, SURUM)
    h(ws, 5, 4, "Lisans")
    h(ws, 5, 5, "Tek kullanıcı")
    h(ws, 7, 1, "KARAR", kalin=True)
    h(ws, 7, 2, "=PANO!B4", kalin=True, boyut=14)
    ozet = [
        (9, "Eşleşme Oranı", "=hih_eslesmeOrani", YÜZDE),
        (10, "Net Beklenen", "=SUM(tblVeriA[NetBeklenen])", TL),
        (11, "Mutabakat Farkı", "=hih_mutabakatFarki", TL),
        (12, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")', CATI),
        (13, "Tutar Farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")', CATI),
        (14, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")', CATI),
        (15, "Kova Bütünlük", "=hih_kovaButunluk", None),
        (16, "Madde Atıf", "=hih_maddeAtifMetni", None),
    ]
    for r, ad, form, fmt in ozet:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, sayi=fmt, kalin=True)
    h(ws, 18, 1, "Varsayımlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 19, 1,
      "Tolerans kuruşu AYARLAR'dan gelir. Anahtar UPPER/TRIM/SUBSTITUTE ile normalize edilir. "
      "Bu çıktı karar destektir; kesin mali/hukuki görüş yerine geçmez. "
      "Hizmet yer teslimi ve belge seti tamamlığı tartışmalı/belirsiz alanlardır (yorum A/B).",
      kaydir=True)
    ws.merge_cells("A19:H19")
    h(ws, 21, 1, "Önerilen Aksiyonlar", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 1, "1. Tutar farkı satırlarında AKSIYON düzeltme metnini kullanın.")
    h(ws, 23, 1, "2. TANIMSIZ kök nedenleri elle inceleyin.")
    h(ws, 24, 1, "3. Belge seti ve yer teslimi belgelerini tamamlayın.")
    baski_hazirla(ws, "A1:H25", f"{URUN_AD} | Kanıt")
    genislik(ws, {"A": 28, "B": 40, "C": 14, "D": 12, "E": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Alıcı Ülke Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 1, "Ulke", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, p in enumerate(ALICI_ULKELER, 7):
        _sari(ws, i, 1, p, baslik="Alıcı Ülke", mesaj="Listeye ekleyebilirsiniz")
    h(ws, 3, 3, "Durum Listesi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 6, 3, "Durum", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, d in enumerate(DURUMLAR, 7):
        _sari(ws, i, 3, d, baslik="Durum", mesaj="Durum değeri")
    h(ws, 3, 5, "Banka Listesi", kalin=True)
    h(ws, 6, 5, "Banka", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    for i, b in enumerate(BANKALAR, 7):
        _sari(ws, i, 5, b, baslik="Banka", mesaj="Banka adı")
    h(ws, 3, 7, "EvetHayir", kalin=True)
    h(ws, 6, 7, "Secim", kalin=True, zemin=ORTA_MAVI, yazi="FFFFFF")
    _sari(ws, 7, 7, "EVET", baslik="Seçim", mesaj="EVET/HAYIR")
    _sari(ws, 8, 7, "HAYIR", baslik="Seçim", mesaj="EVET/HAYIR")
    genislik(ws, {"A": 18, "C": 16, "E": 16, "G": 12})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=10)
    h(ws, 3, 1, "Canlı Kontrol Paneli", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    testler = [
        (5, "A satır sayısı", '=COUNTA(tblVeriA[ReferansNo])'),
        (6, "B satır sayısı", '=COUNTA(tblVeriB[ReferansNo])'),
        (7, "Eşleşti", '=COUNTIF(tblEslesme[Kova],"eslesti")'),
        (8, "Tolerans", '=COUNTIF(tblEslesme[Kova],"tolerans_icinde")'),
        (9, "Tutar farkı", '=COUNTIF(tblEslesme[Kova],"tutar_farki")'),
        (10, "Eşleşmedi", '=COUNTIF(tblEslesme[Kova],"eslesmedi")'),
        (11, "Kova toplam", "=B7+B8+B9+B10"),
        (12, "Bütünlük", '=IF(B11=B5,"BUTUN","KAYIP")'),
        (13, "Duplike A", '=COUNTIF(tblAnahtarA[Duplike],"EVET")'),
        (14, "TANIMSIZ", '=COUNTIF(tblKok[KokNeden],"TANIMSIZ")'),
        (15, "Boş karar yolu", '=IF(B5=0,"VERİ YOK","VERİ VAR")'),
        (16, "Kalite skoru",
         '=IF(B5=0,0,ROUND(100*(B7+B8)/MAX(B5,1),0))'),
        (17, "Net A", "=SUM(tblVeriA[NetBeklenen])"),
        (18, "Net B", "=SUM(tblVeriB[Tutar])"),
        (19, "Fark kontrol", "=B17-B18"),
        (20, "Karar panosu", "=PANO!B4"),
        (21, "İndirim oranı", "=hih_indirimOrani"),
    ]
    for r, ad, form in testler:
        h(ws, r, 1, ad)
        h(ws, r, 2, form, kalin=True,
          sayi=TL if r in (17, 18, 19) else (YÜZDE if r == 21 else (CATI if r < 15 else None)))
    genislik(ws, {"A": 24, "B": 18})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Örnek Senaryo Açıklaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1,
      f"Dosyada {DEMO} satır örnek fatura–transfer mutabakatı vardır: eşleşen, tolerans içi, "
      "tutar farkı ve eşleşmeyen kovalar kasıtlı üretilmiştir. Kısa/uzun referans senaryosu "
      "satır 3, 7, 11'dedir. VERI_A ve VERI_B'yi temizleyip kendi dökümünüzü yapıştırabilirsiniz.",
      kaydir=True)
    ws.merge_cells("A4:H4")
    h(ws, 6, 1, "Örnek özet (bilgi)")
    h(ws, 7, 1, "Beklenen eşleşen ~24 | tolerans ~3 | tutar farkı ~5 | eşleşmedi ~3")
    h(ws, 9, 1, "Ölçek sözleşmesi")
    h(ws, 9, 2, 50000, sayi=CATI)
    h(ws, 11, 1, "Kural seti")
    h(ws, 11, 2, "HIH-001 … HIH-005")
    genislik(ws, {"A": 70, "B": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", RENK, URUN_AD, son_kolon=14)
    alt_bant(ws, 3, "Tek kaynak parametreler — kaynak ve yürürlük zorunlu (D10)", son_kolon=10)
    basliklar = ["anahtar", "deger", "birim", "kaynak", "yururluk_tarihi", "aciklama"]
    baslik_satiri(ws, 5, basliklar)
    params = [
        ("toleransKurus", 0.05, "TL", "İş kuralı — kuruş toleransı", "01.01.2026", "Eşleşme toleransı"),
        ("raporTarihi", RAPOR_TARIHI, "tarih", "Kullanıcı / dönem kapanışı", "01.01.2026", "Rapor tarihi"),
        ("hesapYili", 2026, "yil", "GIRDI bağlantısı", "01.01.2026", "Aktif hesap yılı"),
        ("hih_yilTaban", 2024, "yil", "CHOOSE indeksi", "01.01.2026", "2025→1 için taban"),
        ("hih_esikEslesme", 0.90, "oran", "İç politika / HIH-004", "01.01.2026", "Eşleşme oranı eşiği"),
        ("hih_esikFark", 100, "TL", "İç politika", "01.01.2026", "Fark tutarı eşiği"),
        ("hih_azamiTutar", 10000000, "TL", "İç politika — uç değer", "01.01.2026", "Azami makul tutar"),
        ("hih_ikincilUzunluk", 12, "karakter", "Anahtar stratejisi SPEC", "01.01.2026", "İkincil anahtar uzunluğu"),
        ("hih_olcekHedef", 50000, "satir", "Manda A2 Ö1", "01.01.2026", "Ölçek sözleşmesi"),
        ("hih_anomaliZ", 2.0, "z", "İstatistik kuralı", "01.01.2026", "Anomali Z eşiği"),
        ("hih_yuzdelikOran", 0.9, "oran", "İstatistik kuralı", "01.01.2026", "PERCENTILE oranı"),
        ("hih_senaryoIyi", 1.05, "oran", "Senaryo", "01.01.2026", "İyimser çarpan"),
        ("hih_senaryoKotu", 0.95, "oran", "Senaryo", "01.01.2026", "Kötümser çarpan"),
        ("hih_bazCarpan", 1.0, "oran", "Senaryo", "01.01.2026", "Baz çarpan"),
        ("hih_sifirCarpan", 0.0, "oran", "Senaryo", "01.01.2026", "Nötr çarpan"),
        ("hih_oranArtisPuan", 0.05, "oran", "M-SEN", "01.01.2026", "Kural değişim puanı"),
        ("hih_tahminUst", 1.10, "oran", "Tahmin", "01.01.2026", "Üst bant"),
        ("hih_tahminAlt", 0.90, "oran", "Tahmin", "01.01.2026", "Alt bant"),
        ("hih_riskDusuk", 20, "puan", "Risk", "01.01.2026", "Düşük risk"),
        ("hih_riskOrta", 55, "puan", "Risk", "01.01.2026", "Orta risk"),
        ("hih_riskYuksek", 85, "puan", "Risk", "01.01.2026", "Yüksek risk"),
        ("hih_fiyatModeli", "Tek seferlik alım — aylık ücret yok", "metin", "Ürün konumlandırma", "01.01.2026", "R3 ayrım"),
        ("hih_kvkkBeyani", "Veri cihazdan çıkmaz", "metin", "KVKK beyanı", "01.01.2026", "R3 ayrım"),
        ("hih_normAnahtar", "UPPER TRIM SUBSTITUTE", "metin", "Ç7 anahtar", "01.01.2026", "Serbest alternatif"),
        ("hih_tanimsizKova", "TANIMSIZ görünür", "metin", "E04", "01.01.2026", "Serbest alternatif"),
        ("hih_dosyaSurumu", SURUM, "metin", "Sürüm yönetimi", "01.01.2026", "Ürün sürümü"),
        ("hih_firmaUnvan", "Örnek Hizmet A.Ş.", "metin", "Kullanıcı girişi", "01.01.2026", "Firma"),
        ("hih_parmakIzi", f"HIH-{SURUM}-20260810", "metin", "Parmak izi", "01.01.2026", "Dosya parmak izi"),
        ("vaka_matrah_1", 100000, "TL", "VAKALAR", "01.01.2026", "Vaka 1 matrah"),
        ("vaka_matrah_2", 50000, "TL", "VAKALAR", "01.01.2026", "Vaka 2 matrah"),
        ("vaka_matrah_3", 75000, "TL", "VAKALAR", "01.01.2026", "Vaka 3 matrah"),
    ]
    for i, (ana, deg, bir, kay, yur, acik) in enumerate(params):
        r = 6 + i
        h(ws, r, 1, ana)
        if ana == "hesapYili":
            # GIRDI SSOT — ayna formül kilitli, sarı değil (G04)
            h(ws, r, 2, "=GIRDI!B7", sayi=CATI, yazi=DIS_REF_YEŞIL)
        elif isinstance(deg, date):
            _sari(ws, r, 2, deg, sayi=TARİH, baslik=ana, mesaj=acik)
        elif isinstance(deg, float) and deg <= 1.5:
            _sari(ws, r, 2, deg, sayi=YÜZDE, baslik=ana, mesaj=acik)
        elif isinstance(deg, (int, float)):
            _sari(ws, r, 2, deg, sayi=TL if bir == "TL" else CATI, baslik=ana, mesaj=acik)
        else:
            _sari(ws, r, 2, deg, baslik=ana, mesaj=acik)
        h(ws, r, 3, bir)
        h(ws, r, 4, kay)
        h(ws, r, 5, yur)
        h(ws, r, 6, acik, kaydir=True)

    son_param = 5 + len(params)
    h(ws, son_param + 2, 8, "Motor Cikti", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, son_param + 2, 9, "Deger", kalin=True, yazi=KOYU_LACIVERT)
    motor = [
        (son_param + 3, "hih_mutabakatFarki",
         '=IFERROR(SUM(tblVeriA[NetBeklenen])-SUM(tblVeriB[Tutar]),0)'),
        (son_param + 4, "hih_eslesmeOrani",
         '=IF(COUNTA(tblEslesme[A_Norm])=0,0,'
         '(COUNTIF(tblEslesme[Kova],"eslesti")+COUNTIF(tblEslesme[Kova],"tolerans_icinde"))'
         '/COUNTA(tblEslesme[A_Norm]))'),
        (son_param + 5, "hih_kovaButunluk", "=ESLESTIRME!C2"),
        (son_param + 6, "hih_kokNedenDagilim",
         '=IFERROR(SUMPRODUCT((tblKok[KokNeden]<>"")*1),0)'),
        (son_param + 7, "hih_paretoFark",
         '=IFERROR(IF(SUM(tblEslesme[AbsFark])=0,0,MAX(tblEslesme[AbsFark])/SUM(tblEslesme[AbsFark])),0)'),
        (son_param + 8, "hih_anomaliSayisi",
         '=COUNTIF(tblEslesme[EsikUstu],1)'),
        (son_param + 9, "hih_tahminAralik",
         '=IFERROR(PERCENTILE(tblVeriA[NetBeklenen],hih_yuzdelikOran),0)'),
        (son_param + 10, "hih_tornadoZirve",
         '=IFERROR(MAX(PANO!K27:K31),0)'),
        (son_param + 11, "hih_senaryoKarsilastirma",
         '=IFERROR(SENARYO!B11,0)'),
        (son_param + 12, "hih_kuralDegisimSenaryo",
         '=IFERROR(MOTOR!G41,0)'),
        (son_param + 13, "hih_indirimOrani",
         '=IFERROR(INDEX(tblKurallar[aktif_deger],MATCH("HIH-001",tblKurallar[kural_id],0)),0)'),
        (son_param + 14, "hih_maddeAtifMetni",
         '=IFERROR(MOTOR!G48,"")'),
        (son_param + 15, "hih_kararGerekce",
         '=IFERROR(MOTOR!G34,"")'),
        (son_param + 16, "hih_vakaDurum",
         '=IFERROR(COUNTIF(tblVakalar[durum],"TUTARLI"),0)'),
    ]
    for r, ad, form in motor:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kalin=True,
          sayi=YÜZDE if "Orani" in ad or "pareto" in ad or "indirim" in ad.lower() else TL)

    y0 = son_param + 18
    h(ws, y0, 8, "Canli Yorumlar", kalin=True, yazi=KOYU_LACIVERT)
    yorum_satir = [
        (y0 + 1, "hih_yorumFark",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Fark",TEXT(hih_mutabakatFarki,"0.00"),"TL"),"-")'),
        (y0 + 2, "hih_yorumKokNeden",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kok neden sinifi",TEXT(hih_kokNedenDagilim,"0")),"-")'),
        (y0 + 3, "hih_yorumTornado",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tornado",TEXT(hih_tornadoZirve,"0.00")),"-")'),
        (y0 + 4, "hih_yorumTahmin",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Tahmin",TEXT(hih_tahminAralik,"0.00")),"-")'),
        (y0 + 5, "hih_yorumSenaryo",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Senaryo bant",TEXT(hih_senaryoKarsilastirma,"0.00")),"-")'),
        (y0 + 6, "hih_yorumKuralDegisim",
         '=IFERROR(_xlfn.TEXTJOIN(" ",TRUE,"Kural degisim farki",TEXT(hih_kuralDegisimSenaryo,"0.00")),"-")'),
    ]
    for r, ad, form in yorum_satir:
        h(ws, r, 8, ad)
        h(ws, r, 9, form, kaydir=True)

    # Store row maps for adlari_bagla
    ws._hih_motor_rows = {ad: r for r, ad, _ in motor}  # type: ignore[attr-defined]
    ws._hih_yorum_rows = {ad: r for r, ad, _ in yorum_satir}  # type: ignore[attr-defined]
    ws._hih_param_count = len(params)  # type: ignore[attr-defined]

    genislik(ws, {"A": 28, "B": 42, "C": 12, "D": 28, "E": 14, "F": 28, "H": 28, "I": 42})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=12)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    adimlar = [
        "GIRDI: Dönem, belge tamamlık ve yer teslimi riskini girin.",
        "VERI_A: Hizmet ihracatı faturalarını sarı hücrelere yapıştırın.",
        "VERI_B: Banka transferlerini girin.",
        "ANAHTAR / ESLESTIRME: Normalize ve dört kova otomatik çalışır.",
        "KURALLAR / MOTOR: HIH-001…005 kural yılı ve karar zinciri.",
        "KOK_NEDEN / AKSIYON: Tutar farkı kök nedeni ve düzeltme metni.",
        "KANIT_RAPORU: Dönem dosyasına PDF olarak yazdırın.",
    ]
    for i, m in enumerate(adimlar, 5):
        h(ws, i, 1, f"{i - 4}. {m}", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)
    h(ws, 13, 1, "Belirsizlik beyanı", kalin=True, yazi="B3261E")
    h(ws, 14, 1,
      "Hizmet yer teslimi ve belge seti tamamlığı tartışmalı/belirsiz alanlardır. "
      "Yorum A: sıkı belge seti; Yorum B: esnek tamamlık. Dosya her iki yorumu SENARYO'da gösterir.",
      kaydir=True)
    ws.merge_cells("A14:J14")
    h(ws, 16, 1, "Sık yapılan hatalar", kalin=True, yazi="B3261E")
    for i, m in enumerate([
        "Referans numarasını boşluklu bırakmak (normalize edilir ama kontrol edin)",
        "A ve B dönemlerini karıştırmak",
        "Net beklenen yerine brüt faturayı B tarafına yazmak",
    ], 17):
        h(ws, i, 1, "• " + m)
    h(ws, 21, 1, f"Sürüm {SURUM} | Bu dosya karar destek aracıdır.", yazi=GRİ)
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
                  baslik="Ülke", mesaj="Ad girin",
                  hata_baslik="Ad", hata_mesaj="En az 1 karakter",
                  isaret="greaterThan", f2="40")
    for r in range(7, 11):
        dogrulama(ws, "textLength", "1", f"C{r}",
                  baslik="Durum", mesaj="Durum girin",
                  hata_baslik="Durum", hata_mesaj="Geçersiz",
                  isaret="greaterThan", f2="30")
    ws = wb["KANIT_RAPORU"]
    for r, ad in [(5, "Sürüm"), (22, "Aksiyon1"), (23, "Aksiyon2"), (24, "Aksiyon3")]:
        dogrulama(ws, "textLength", "0", f"B{r}" if r == 5 else f"A{r}",
                  baslik=ad, mesaj="Metin alanı",
                  hata_baslik=ad, hata_mesaj="Geçersiz",
                  isaret="greaterThanOrEqual", f2="200")


def kosullu_bicim(wb):
    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    sari = PatternFill("solid", fgColor="FFEB9C")
    yf = Font(color="006100", name=FONT)
    kf = Font(color="9C0006", name=FONT)

    ws = wb["PANO"]
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"UYGUN DEĞİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B4", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))
    for r in range(6, 24):
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws.conditional_formatting.add(f"B{r}", CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))

    ws = wb["ESLESTIRME"]
    for col in ("G",):
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"eslesti"'], fill=yesil, font=yf))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"tutar_farki"'], fill=sari))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"eslesmedi"'], fill=kirmizi, font=kf))
        ws.conditional_formatting.add(f"{col}{ILK}:{col}{SON}",
            CellIsRule(operator="equal", formula=['"tolerans_icinde"'], fill=sari))

    ws = wb["KOK_NEDEN"]
    ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
        CellIsRule(operator="equal", formula=['"TANIMSIZ"'], fill=kirmizi, font=kf))
    for kod in ("STOPAJ", "KUR_FARKI", "KISMI_TRANSFER", "BELGE_EKSIK"):
        ws.conditional_formatting.add(f"H16:H{15 + KAPASITE}",
            CellIsRule(operator="equal", formula=[f'"{kod}"'], fill=sari))

    ws = wb["AKSIYON"]
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"YUKSEK"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"E{ILK}:E{SON}",
        CellIsRule(operator="equal", formula=['"NORMAL"'], fill=sari))

    ws = wb["KANIT_RAPORU"]
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"UYGUN DEĞİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("B7", CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["VERI_A"]
    ws.conditional_formatting.add(f"H{ILK}:H{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))

    ws = wb["VERI_B"]
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"C{ILK}:C{SON}",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=sari))

    ws = wb["ANAHTAR"]
    ws.conditional_formatting.add(f"F{ILK}:F{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add(f"N{ILK}:N{SON}",
        CellIsRule(operator="equal", formula=['"EVET"'], fill=kirmizi, font=kf))

    ws = wb["VAKALAR"]
    ws.conditional_formatting.add("G6:G9",
        CellIsRule(operator="equal", formula=['"TUTARLI"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("G6:G9",
        CellIsRule(operator="equal", formula=['"KIRIK"'], fill=kirmizi, font=kf))

    ws = wb["MOTOR"]
    ws.conditional_formatting.add("G33",
        CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
    ws.conditional_formatting.add("G33",
        CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
    ws.conditional_formatting.add("G33",
        CellIsRule(operator="equal", formula=['"UYGUN DEĞİL"'], fill=kirmizi, font=kf))
    ws.conditional_formatting.add("G33",
        CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["SENARYO"]
    for r in range(6, 10):
        ws.conditional_formatting.add(f"E{r}",
            CellIsRule(operator="equal", formula=['"UYGUN"'], fill=yesil, font=yf))
        ws.conditional_formatting.add(f"E{r}",
            CellIsRule(operator="equal", formula=['"DİKKAT"'], fill=sari))
        ws.conditional_formatting.add(f"E{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=kirmizi, font=kf))

    ws = wb["GIRDI"]
    ws.conditional_formatting.add("B9",
        CellIsRule(operator="lessThan", formula=["0.95"], fill=sari))
    ws.conditional_formatting.add("B10",
        CellIsRule(operator="greaterThan", formula=["0.3"], fill=kirmizi, font=kf))

    if "KONTROLLER" in wb.sheetnames:
        ws = wb["KONTROLLER"]
        ws.conditional_formatting.add("B12", CellIsRule(operator="equal", formula=['"BUTUN"'], fill=yesil, font=yf))
        ws.conditional_formatting.add("B12", CellIsRule(operator="equal", formula=['"KAYIP"'], fill=kirmizi, font=kf))
        ws.conditional_formatting.add("B16", CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yf))
        ws.conditional_formatting.add("B16", CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kf))


def adlari_bagla(wb):
    params = [
        "toleransKurus", "raporTarihi", "hesapYili", "hih_yilTaban",
        "hih_esikEslesme", "hih_esikFark", "hih_azamiTutar",
        "hih_ikincilUzunluk", "hih_olcekHedef", "hih_anomaliZ", "hih_yuzdelikOran",
        "hih_senaryoIyi", "hih_senaryoKotu", "hih_bazCarpan", "hih_sifirCarpan", "hih_oranArtisPuan",
        "hih_tahminUst", "hih_tahminAlt",
        "hih_riskDusuk", "hih_riskOrta", "hih_riskYuksek",
        "hih_fiyatModeli", "hih_kvkkBeyani", "hih_normAnahtar", "hih_tanimsizKova",
        "hih_dosyaSurumu", "hih_firmaUnvan", "hih_parmakIzi",
        "vaka_matrah_1", "vaka_matrah_2", "vaka_matrah_3",
    ]
    for i, ana in enumerate(params):
        if ana == "hesapYili":
            ad_ekle(wb, ana, "GIRDI!$B$7")
        else:
            ad_ekle(wb, ana, f"AYARLAR!$B${6 + i}")

    # AYARLAR B8 yalnızca görüntü — GIRDI SSOT; sarı sabit kalır, formül yazılmaz
    ws = wb["AYARLAR"]
    motor_rows = getattr(ws, "_hih_motor_rows", {})
    yorum_rows = getattr(ws, "_hih_yorum_rows", {})
    for ad, r in motor_rows.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")
    for ad, r in yorum_rows.items():
        ad_ekle(wb, ad, f"AYARLAR!$I${r}")

    ad_ekle(wb, "ListeAliciUlke", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeDurumlar", "LISTELER!$C$7:$C$10")
    ad_ekle(wb, "ListeBankalar", "LISTELER!$E$7:$E$11")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$G$7:$G$8")


def main(cikti_yolu=None):
    wb = Workbook()
    # SPEC sayfa_akisi + KAPAK önde; AYARLAR/KILAVUZ sonda
    siralar = [
        (kapak, "KAPAK"),
        (pano, "PANO"),
        (girdi, "GIRDI"),
        (veri_a, "VERI_A"),
        (veri_b, "VERI_B"),
        (anahtar, "ANAHTAR"),
        (eslestirme, "ESLESTIRME"),
        (kok_neden, "KOK_NEDEN"),
        (kurallar, "KURALLAR"),
        (motor, "MOTOR"),
        (senaryo, "SENARYO"),
        (vakalar, "VAKALAR"),
        (aksiyon, "AKSIYON"),
        (kanit_raporu, "KANIT_RAPORU"),
        (kontroller, "KONTROLLER"),
        (ornek_veri, "ORNEK_VERI"),
        (listeler, "LISTELER"),
        (ayarlar, "AYARLAR"),
        (kilavuz, "KILAVUZ"),
    ]
    ilk = wb.active
    for i, (fn, ad) in enumerate(siralar):
        ws = ilk if i == 0 else wb.create_sheet()
        fn(ws)

    adlari_bagla(wb)
    _cok_dogrulama(wb)
    kosullu_bicim(wb)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=ILK, satir_sonu=SON)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=16, satir_sonu=15 + KAPASITE)
    # VAKALAR / KURALLAR küçük tablolar
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=11)
    tablo_formullerini_hucrelere_yaz(wb, satir_basi=17, satir_sonu=24)

    for ws in wb.worksheets:
        sayfa_koru(ws)

    wb.calculation.fullCalcOnLoad = True
    urun_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    dosya_ad = "HizmetIhracati100IndirimTransferTakip.xlsx"
    hedef = cikti_yolu or os.path.join(urun_dir, dosya_ad)
    os.makedirs(os.path.dirname(hedef) or ".", exist_ok=True)
    wb.save(hedef)

    cikti = os.path.join(KOK, "cikti", dosya_ad)
    os.makedirs(os.path.dirname(cikti) or ".", exist_ok=True)
    if os.path.abspath(hedef) != os.path.abspath(cikti):
        import shutil
        shutil.copy2(hedef, cikti)

    print(f"Dosya oluşturuldu: {hedef}")
    print(f"Kopya: {cikti}")
    return hedef


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
