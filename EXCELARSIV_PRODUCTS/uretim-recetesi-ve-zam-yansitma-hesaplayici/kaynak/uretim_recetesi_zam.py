"""
Üretim Reçetesi & Zam Yansıtma Hesaplayıcı — üretim betiği.
Manda v4 uyumlu: sayfa sırası, sarı giriş, 1234 koruma, ipucu/not, tablo mimarisi.
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from excel_uretim.ortak import (
    ACIK_GRI,
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
    giris_hucreleri,
    h,
    sabitle,
    sayfa_hazirla,
    sayfa_koru,
    tablo_ekle,
    tablo_formullerini_hucrelere_yaz,
    yorum_ekle,
)

URUN_AD = "Üretim Reçetesi & Zam Yansıtma Hesaplayıcı"
SURUM = "1.0.0"
RENK = "2E7D32"

BIRIMLER = ["kg", "gr", "lt", "adet", "paket", "m"]
HAMMADDELER = ["Un", "Tam Buğday Unu", "Margarin", "Yumurta", "Tuz", "Şeker", "Maya", "Susam", "Diğer"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Reçetedeki hammadde maliyetlerinizi girin; hammadde zamlarını yansıtarak "
                "yeni reçete maliyetini, fire dahil birim maliyeti ve hedef kâr marjını "
                "koruyan yeni satış fiyatını gerekçesiyle hesaplayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Ürün bazında hammadde reçetesi; birim, miktar, fire ve hammadde zamlarıyla yeniden maliyetleme",
        "Hammadde zamlarının reçete maliyetine, kâr marjına ve satış fiyatına toplam etkisinin hesaplanması",
        "Hedef kâr marjını koruyan önerilen yeni satış fiyatının ve yansıtılacak zam oranının belirlenmesi",
        "Hammadde maliyet payı yoğunlaşması, anomali ve veri kalite skoru ile karar kapısı ve aksiyon önerisi",
        "Zam senaryosu bant aralığı ve hammadde duyarlılık (tornado) analizi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Hammadde zamlarını ürün fiyatlarına yansıtmak zorunda kalan üretim işletmeleri",
        "Reçete maliyetini fire ve işçilik dahil takip eden gıda/üretim firmaları",
        "Fiyat güncelleme kararını yönetime kanıtla sunmak isteyen satın alma ve finans yöneticileri",
        "Müşterisine maliyet ve zam yansıtma analizi öneren mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 20, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. RECETE_LISTESI sayfasına reçete kalemlerinizi girin. "
                 "3. PANO ve KARAR sayfalarından zam yansıtma kararını izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A21:P21")
    h(ws, 23, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A23:P23")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Reçete kalemlerini girin",
         "RECETE_LISTESI sayfasındaki sarı hücrelere ürün kodu/adı, hammadde, birim, "
         "miktar, birim fiyat, fire ve zam oranlarını girin. Tutarlar otomatik hesaplanır."),
        ("Adım 2 — Zamları yansıtın",
         "Hammadde zam oranlarını girin; reçete maliyeti, fire dahil maliyet ve önerilen "
         "satış fiyatı anında yeniden hesaplanır."),
        ("Adım 3 — Kararı değerlendirin",
         "KARAR sayfası zam yansıtma kararını (ZAM YANSIT / KISMİ YANSIT / FİYATI KORU) "
         "gerekçesiyle üretir."),
        ("Adım 4 — Oranları ayarlayın",
         "Hedef kâr marjı, işçilik ve genel üretim gideri oranları AYARLAR sayfasındadır."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Zam oranını yüzde olarak 15 değil 0,15 girmek",
        "Fire oranını toplama ekleyip birim fiyata uygulamak (fire, miktara uygulanır)",
        "Miktar girilmeden birim fiyat girmek (tutar sıfır görünür)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 16, 1, "Sık sorulan sorular", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sss = [
        "Soru: Zam oranını nasıl girerim? Cevap: 0,15 olarak (onluk kesir). %15 yazmak hesabı 100 kat büyütür.",
        "Soru: Fire oranı nereye uygulanır? Cevap: Fire, reçetede kullanılan hammadde miktarına eklenir; tutar miktar üzerinden hesaplanır.",
        "Soru: Önerilen satış fiyatı nasıl bulunur? Cevap: Zam sonrası birim maliyet, hedef kâr marjına göre (maliyet / (1 - marj)) fiyatlanır.",
        "Soru: Hammadde listesinde aradığım yoksa? Cevap: LISTELER sayfasındaki sarı hücreye yeni hammadde adını ekleyin.",
    ]
    for i, m in enumerate(sss, 17):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
        ws.row_dimensions[i].height = 32
    h(ws, 22, 1, "Bu dosya karar destek aracıdır; maliyet ve fiyatlama hesabı mali müşavir görüşü "
                 "yerine geçmez.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def recete_listesi(ws):
    sayfa_hazirla(ws, "RECETE_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Üretim Reçetesi Kalemleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Hammadde tutarı, fire dahil tutar, zam sonrası "
                "tutar ve fark otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ürün Kodu", "Ürün Adı", "Hammadde", "Birim", "Miktar",
                "Birim Fiyat (TL)", "Fire Oranı", "Zam Oranı",
                "Hammadde Tutarı (TL)", "Fire Dahil Tutar (TL)",
                "Zam Sonrası Birim Fiyat (TL)", "Zam Sonrası Tutar (TL)",
                "Fark (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Hammadde Tutarı (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="",0,tblRecete[[#This Row],[Miktar]]*tblRecete[[#This Row],[Birim Fiyat (TL)]])',
        "Fire Dahil Tutar (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="",0,tblRecete[[#This Row],[Miktar]]*(1+tblRecete[[#This Row],[Fire Oranı]])*tblRecete[[#This Row],[Birim Fiyat (TL)]])',
        "Zam Sonrası Birim Fiyat (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="",0,tblRecete[[#This Row],[Birim Fiyat (TL)]]*(1+tblRecete[[#This Row],[Zam Oranı]]))',
        "Zam Sonrası Tutar (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="",0,tblRecete[[#This Row],[Miktar]]*(1+tblRecete[[#This Row],[Fire Oranı]])*tblRecete[[#This Row],[Zam Sonrası Birim Fiyat (TL)]])',
        "Fark (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="",0,tblRecete[[#This Row],[Zam Sonrası Tutar (TL)]]-tblRecete[[#This Row],[Fire Dahil Tutar (TL)]])',
        "Açıklama": '=IF(tblRecete[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblRecete[[#This Row],[Ürün Adı]]," | Hammadde: ",tblRecete[[#This Row],[Hammadde]]," | Fark: ",TEXT(tblRecete[[#This Row],[Fark (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblRecete", "A5:N1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeHammaddeler", "C6:C1004",
              baslik="Hammadde", mesaj="Listeden hammaddeyi seçin.",
              hata_baslik="Geçersiz hammadde", hata_mesaj="Listede olmayan hammadde giremezsiniz.")
    dogrulama(ws, "list", "ListeBirimler", "D6:D1004",
              baslik="Birim", mesaj="Listeden birimi seçin.",
              hata_baslik="Geçersiz birim", hata_mesaj="Listede olmayan birim giremezsiniz.")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Miktar", mesaj="Reçetedeki hammadde miktarını girin.",
              hata_baslik="Geçersiz miktar", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Birim Fiyat", mesaj="Birim fiyatı TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000 arasında olmalı.",
              isaret="between", f2="1000000")
    for kolon in ("G", "H"):
        dogrulama(ws, "decimal", "0", f"{kolon}6:{kolon}1004",
                  baslik="Oran", mesaj="Oranı 0-1 arasında girin (örn. %10 için 0,10).",
                  hata_baslik="Geçersiz oran", hata_mesaj="0 ile 1 arasında olmalı.",
                  isaret="between", f2="1")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=5).number_format = CATI
        for kolon in (6, 9, 10, 11, 12, 13):
            ws.cell(row=satir, column=kolon).number_format = TL
        for kolon in (7, 8):
            ws.cell(row=satir, column=kolon).number_format = YÜZDE
    ornek = [
        ("UR-01", "Sade Poğaça", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-01", "Sade Poğaça", "Margarin", "kg", 0.020, 90, 0.03, 0.08),
        ("UR-01", "Sade Poğaça", "Yumurta", "kg", 0.020, 45, 0.05, 0.05),
        ("UR-01", "Sade Poğaça", "Tuz", "kg", 0.002, 8, 0.02, 0.02),
        ("UR-01", "Sade Poğaça", "Maya", "kg", 0.003, 30, 0.03, 0.04),
        ("UR-02", "Tam Buğday Ekmeği", "Un", "kg", 0.060, 20, 0.05, 0.15),
        ("UR-02", "Tam Buğday Ekmeği", "Tam Buğday Unu", "kg", 0.020, 25, 0.05, 0.12),
        ("UR-02", "Tam Buğday Ekmeği", "Maya", "kg", 0.004, 30, 0.03, 0.04),
        ("UR-02", "Tam Buğday Ekmeği", "Tuz", "kg", 0.003, 8, 0.02, 0.02),
        ("UR-03", "Açma", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-03", "Açma", "Margarin", "kg", 0.030, 90, 0.03, 0.08),
        ("UR-03", "Açma", "Yumurta", "kg", 0.030, 45, 0.05, 0.05),
        ("UR-03", "Açma", "Şeker", "kg", 0.005, 30, 0.02, 0.03),
        ("UR-03", "Açma", "Maya", "kg", 0.003, 30, 0.03, 0.04),
        ("UR-04", "Simit", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-04", "Simit", "Susam", "kg", 0.020, 60, 0.06, 0.06),
        ("UR-04", "Simit", "Maya", "kg", 0.003, 30, 0.03, 0.04),
        ("UR-04", "Simit", "Şeker", "kg", 0.005, 30, 0.02, 0.03),
    ]
    for i, (kod, urun, ham, birim, miktar, fiyat, fire, zam) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = kod
        ws.cell(row=i, column=2).value = urun
        ws.cell(row=i, column=3).value = ham
        ws.cell(row=i, column=4).value = birim
        ws.cell(row=i, column=5).value = miktar
        ws.cell(row=i, column=6).value = fiyat
        ws.cell(row=i, column=7).value = fire
        ws.cell(row=i, column=8).value = zam
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8])
    sabitle(ws, "A6")
    genislik(ws, {"A": 10, "B": 20, "C": 18, "D": 8, "E": 9, "F": 13, "G": 11,
                  "H": 11, "I": 16, "J": 18, "K": 20, "L": 18, "M": 13, "N": 52})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Maliyet ve Zam Yansıtma Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Reçete kalemlerinden otomatik beslenir; ürün bazında eski/yeni maliyet, "
                "artış ve önerilen fiyat canlı üretilir.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ürün Kodu", "Ürün Adı", "Kalem Sayısı", "Eski Reçete Maliyeti (TL)",
                "Fire Dahil Eski Maliyet (TL)", "Zam Sonrası Maliyet (TL)",
                "Maliyet Artışı (TL)", "Maliyet Artış Oranı", "Hedef Fiyat (TL)",
                "Fark (TL)"]
    baslik_satiri(ws, 5, sutunlar)
    urun_kodlari = ["UR-01", "UR-02", "UR-03", "UR-04"]
    for r, kod in enumerate(urun_kodlari, 6):
        h(ws, r, 1, kod, kalin=True, yazi="333333")
        h(ws, r, 2, f'=IFERROR(INDEX(tblRecete[Ürün Adı],MATCH($A{r},tblRecete[Ürün Kodu],0)),"")',
          yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=COUNTIF(tblRecete[Ürün Kodu],$A{r})', sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblRecete[Hammadde Tutarı (TL)],tblRecete[Ürün Kodu],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblRecete[Fire Dahil Tutar (TL)],tblRecete[Ürün Kodu],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=SUMIFS(tblRecete[Zam Sonrası Tutar (TL)],tblRecete[Ürün Kodu],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f'=F{r}-E{r}', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 8, f'=IFERROR(G{r}/E{r},0)', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, r, 9, f'=IFERROR((F{r}+E{r}*UreticiIscilikOran+F{r}*UreticiGugOran)/(1-UreticiHedefMarjOran),0)',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 10, f'=I{r}-F{r}', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "TOPLAM BLOK", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    toplamlar = [
        (13, "Toplam Hammadde Maliyeti (TL)", "=SUM(D6:D9)", TL),
        (14, "Toplam Fire Dahil Maliyet (TL)", "=SUM(E6:E9)", TL),
        (15, "Toplam Zam Sonrası Maliyet (TL)", "=SUM(F6:F9)", TL),
        (16, "TOPLAM MALİYET ARTIŞI (TL)", "=SUM(G6:G9)", TL),
        (17, "Ortalama Maliyet Artış Oranı", "=IFERROR(B16/B14,0)", YÜZDE),
        (18, "Toplam İşçilik Maliyeti (TL)", "=SUM(E6:E9)*UreticiIscilikOran", TL),
        (19, "Toplam GÜG Maliyeti (TL)", "=SUM(F6:F9)*UreticiGugOran", TL),
        (20, "Genel Maliyet Artış Oranı", "=IFERROR(B16/B14,0)", YÜZDE),
        (21, "Zam Yansıtma Gereksinimi", '=IF(B20>UreticiZamYansitmaEsik,"GEREKLİ","GEREKSİZ")', None),
        (22, "En Çok Etkilenen Ürün", '=IFERROR(INDEX(A6:A9,MATCH(MAX(G6:G9),G6:G9,0)),"")', None),
        (23, "En Yüksek Maliyet Artışı (TL)", "=IFERROR(MAX(G6:G9),0)", TL),
    ]
    for satir, ad, form, sayi in toplamlar:
        kalin = satir in (16,)
        h(ws, satir, 1, ad, yazi="333333", kalin=kalin)
        h(ws, satir, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=kalin)
    # Panel verisi (grafikler için)
    h(ws, 24, 1, "Ürün Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 24, 2, "Eski Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 24, 4, "Yeni Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 24, 6, "Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 25 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 25 + i, 2, f"=IFERROR(INDEX(D6:D9,UreticiPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 25 + i, 4, f"=IFERROR(INDEX(F6:F9,UreticiPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 25 + i, 6, f"=IFERROR(INDEX(G6:G9,UreticiPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Projeksiyon
    h(ws, 38, 1, "Ortalama Maliyet Artışı (TL)", yazi="333333")
    h(ws, 38, 2, "=IFERROR(AVERAGE(G6:G9),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 39, 1, "Gelecek Dönem Maliyet Artışı Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 39, 2, "=IFERROR(FORECAST.LINEAR(UreticiPanelSatir_12,G6:G9,UretimProjX)-G9,0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 40, 1, "Maliyet Artışı Std Sapması (TL)", yazi="333333")
    h(ws, 40, 2, "=IFERROR(STDEV.P(G6:G9),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 41, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 41, 2, "=IFERROR(B38-B40*UreticiTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 42, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 42, 2, "=IFERROR(B38+B40*UreticiTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Projeksiyon yardımcı serisi (dönem / maliyet artışı)
    h(ws, 45, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 45, 2, "Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for k in range(4):
        h(ws, 46 + k, 1, k + 1, sayi=CATI, yazi=GRİ)
        h(ws, 46 + k, 2, f"=G{6+k}", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 44, "Motor, reçete kalemlerinden canlı beslenir; tüm oranlar AYARLAR'dan okunur.")
    genislik(ws, {"A": 34, "B": 40, "C": 22, "D": 22, "E": 24, "F": 24, "G": 22, "H": 22})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Veri Kalite ve Kontrol Merkezi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Girdilerin kalitesini ölçer; eksik, negatif ve uç değerleri raporlar.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    baslik_satiri(ws, 5, ["Kontrol", "Sonuç", "Durum"])
    kontroller = [
        ("Reçete kalemi girildi mi?", '=IF(COUNTA(tblRecete[Ürün Adı])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Negatif birim fiyat var mı?", '=IF(COUNTIF(tblRecete[Birim Fiyat (TL)],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, "B3261E"),
        ("Negatif miktar var mı?", '=IF(COUNTIF(tblRecete[Miktar],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, "B3261E"),
        ("Fire oranı 0,50 üzerinde mi?", '=IF(COUNTIF(tblRecete[Fire Oranı],">0.5")=0,"TEMİZ","FİRE >0,5 VAR")', NORMAL, "B3261E"),
        ("Zam oranı 0,50 üzerinde mi?", '=IF(COUNTIF(tblRecete[Zam Oranı],">0.5")=0,"TEMİZ","ZAM >0,5 VAR")', NORMAL, "B3261E"),
        ("Zam yansıtma gereksinimi", '=IF(HESAP!B21="GEREKLİ","VAR","YOK")', "B3261E", NORMAL),
        ("Ürün sayısı yeterli mi?", '=IF(COUNTIF(tblRecete[Ürün Kodu],"<>")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
    ]
    for i, (ad, form, durum1, durum2) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "AUTOMATİK", yazi=GRİ)
    h(ws, 14, 1, "Veri Kalite Skoru", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Doluluk Oranı (%)", yazi="333333")
    h(ws, 15, 2, '=IFERROR(COUNTA(tblRecete[Ürün Adı])/(COUNTA(tblRecete[Ürün Adı])+COUNTBLANK(tblRecete[Ürün Adı]))*100,0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Veri Kalite Skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, '=MIN(100,MAX(0,B15))', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Kalite Durumu", yazi="333333")
    h(ws, 17, 2, '=IF(B16>=UreticiKaliteIyi,"YÜKSEK",IF(B16>=UreticiKaliteOrta,"ORTA","DÜŞÜK"))', yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Temel Giriş Kolonu Sayısı", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblRecete[Ürün Adı])*UreticiTemelKolonSayisi", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Canlı Formül Sayacı", yazi="333333")
    h(ws, 19, 2, '=IF(COUNTIF(tblRecete[Açıklama],"")=0,COUNTA(tblRecete[Açıklama]),0)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 21, "Tüm kontroller canlı formüllerle üretilir; sahte geçiş yoktur.")
    genislik(ws, {"A": 38, "B": 46, "C": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Zam Yansıtma Karar Kapısı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Maliyet artışı, hedef marj ve fiyat karşılaştırmasına göre karar üretir.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Reçete Kalem Sayısı", yazi="333333")
    h(ws, 6, 2, "=COUNTA(tblRecete[Ürün Adı])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Genel Maliyet Artış Oranı", yazi="333333")
    h(ws, 7, 2, "=HESAP!B20", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Zam Yansıtma Eşiği", yazi="333333")
    h(ws, 8, 2, "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiZamYansitmaEsik,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Toplam Maliyet Artışı (TL)", yazi="333333")
    h(ws, 9, 2, "=HESAP!B16", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Hedef Kâr Marjı", yazi="333333")
    h(ws, 10, 2, "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiHedefMarjOran,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblRecete[Ürün Adı])=0,"VERİ YOK",'
                 'IF(OR(COUNTIF(tblRecete[Birim Fiyat (TL)],"<0")>0,COUNTIF(tblRecete[Miktar],"<0")>0),"DURDUR",'
                 'IF(HESAP!B20>UreticiZamYansitmaEsik,"ZAM YANSIT",'
                 'IF(AND(HESAP!B20>0,HESAP!B20<=UreticiZamYansitmaEsik),"KISMİ YANSIT","FİYATI KORU"))))',
      boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 12, 2, '=IF(B11="VERİ YOK","Reçete kalemi girilmedi; karar üretilemiyor.",'
                 'IF(B11="DURDUR","Negatif birim fiyat veya miktar tespit edildi; veriler düzeltilmeli.",'
                 'IF(B11="ZAM YANSIT","Genel maliyet artışı eşiğin üzerinde; zam yansıtma gereklidir.",'
                 'IF(B11="KISMİ YANSIT","Maliyet artışı eşiğin altında ancak pozitif; kısmi zam yansıtılabilir.",'
                 '"Maliyet artışı yok veya ihmal edilebilir; mevcut fiyat korunabilir."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[12].height = 44
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["Aksiyon 1", "Aksiyon 2", "Aksiyon 3"], 15):
        h(ws, i, 1, ad, yazi="333333", kalin=True)
        h(ws, i, 2, f'=IF(B11="VERİ YOK","Reçete kalemlerini girin; karar otomatik üretilir.",'
                    f'IF(B11="DURDUR","Negatif değerleri düzeltin.",UreticiAksiyon{i-14}_ACIKLAMA))',
          yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[i].height = 40
    h(ws, 19, 1, "Karar Notu", yazi="333333")
    h(ws, 19, 2, "Bu karar destek aracıdır; fiyat kararı yönetimin onayına tabidir.", yazi=GRİ)
    baski_hazirla(ws, "A1:B19", URUN_AD)
    alt_bant(ws, 21, "Karar kuralları KILAVUZ sayfasında gösterilmiştir.")
    genislik(ws, {"A": 30, "B": 80})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (1, "Toplam Hammadde Maliyeti (TL)", "=HESAP!B13", TL),
        (2, "Toplam Zam Sonrası Maliyet (TL)", "=HESAP!B15", TL),
        (3, "TOPLAM MALİYET ARTIŞI (TL)", "=HESAP!B16", TL),
        (4, "Ortalama Maliyet Artış Oranı", "=HESAP!B17", YÜZDE),
        (5, "Zam Yansıtma Gereksinimi", "=HESAP!B21", NORMAL),
        (6, "En Çok Etkilenen Ürün", "=HESAP!B22", NORMAL),
        (7, "Karar", "=KARAR!B11", NORMAL),
        (8, "Kalite Skoru", "=KONTROLLER!B16", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En Çok Etkilenen Ürün", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B22", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "En Yüksek Maliyet Artışı (TL)", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 6, 8, "=HESAP!B23", sayi=TL, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    ws.row_dimensions[6].height = 30
    # Ürün bazlı grafik verisi
    h(ws, 10, 1, "Ürün", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Eski Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Yeni Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Artış Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    urun_etiketleri = ["UR-01 Sade Poğaça", "UR-02 Tam Buğday Ekmeği", "UR-03 Açma", "UR-04 Simit"]
    for i, paz in enumerate(urun_etiketleri, 11):
        h(ws, i, 1, paz, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!D{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!F{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!G{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=HESAP!H{i-5}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Ürün Bazlı Eski ve Yeni Maliyet"
    g1.add_data(Reference(ws, min_col=2, max_col=3, min_row=10, max_row=15), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g1.height, g1.width = 9, 15
    ws.add_chart(g1, "F8")
    g2 = BarChart()
    g2.type = "col"
    g2.title = "Ürün Bazlı Maliyet Artışı"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=15), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g9 = BarChart()
    g9.type = "col"
    g9.title = "Ürün Bazlı Maliyet Artış Oranı"
    g9.add_data(Reference(ws, min_col=5, min_row=10, max_row=15), titles_from_data=True)
    g9.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g9.height, g9.width = 9, 12
    ws.add_chart(g9, "AH8")
    # Hammadde payı dağılımı
    h(ws, 18, 1, "Hammadde", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 18, 2, "Eski Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 18, 3, "Zam Sonrası Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ham in enumerate(["Un", "Tam Buğday Unu", "Margarin", "Yumurta", "Maya", "Şeker", "Susam", "Tuz"], 19):
        h(ws, i, 1, ham, yazi=GRİ)
        h(ws, i, 2, f'=SUMIFS(tblRecete[Hammadde Tutarı (TL)],tblRecete[Hammadde],$A{i})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f'=SUMIFS(tblRecete[Zam Sonrası Tutar (TL)],tblRecete[Hammadde],$A{i})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
    g3 = PieChart()
    g3.title = "Hammadde Bazlı Zam Sonrası Maliyet Payı"
    g3.add_data(Reference(ws, min_col=3, min_row=18, max_row=26), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=19, max_row=26))
    g3.height, g3.width = 9, 12
    ws.add_chart(g3, "F21")
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Hammadde Bazlı Eski ve Zam Sonrası Tutar"
    g5.add_data(Reference(ws, min_col=2, max_col=3, min_row=18, max_row=26), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=19, max_row=26))
    g5.height, g5.width = 9, 12
    ws.add_chart(g5, "AH21")
    # Panel trendi
    h(ws, 28, 1, "Ürün Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 28, 2, "Eski Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 28, 3, "Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 28, 4, "Yeni Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 29 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 29 + i, 2, f"=HESAP!B{25 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 29 + i, 3, f"=HESAP!F{25 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 29 + i, 4, f"=HESAP!D{25 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Ürün Bazlı Maliyet Karşılaştırma"
    g4.add_data(Reference(ws, min_col=2, max_col=4, min_row=28, max_row=40), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=29, max_row=40))
    g4.height, g4.width = 9, 15
    ws.add_chart(g4, "F34")
    # Projeksiyon
    h(ws, 42, 1, "Ortalama Maliyet Artışı (TL)", yazi="333333")
    h(ws, 42, 2, "=HESAP!B38", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 43, 1, "Gelecek Dönem Artış Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 43, 2, "=HESAP!B39", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 44, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 44, 2, "=HESAP!B41", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 45, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 45, 2, "=HESAP!B42", sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = BarChart()
    g6.type = "col"
    g6.title = "Projeksiyon: Artış, Alt ve Üst Sınır"
    g6.add_data(Reference(ws, min_col=2, min_row=42, max_row=45), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=43, max_row=45))
    g6.height, g6.width = 9, 15
    ws.add_chart(g6, "U47")
    # Senaryo verisi
    h(ws, 47, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 47, 6, "Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 48):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-41}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g8 = LineChart()
    g8.title = "Senaryo Maliyet Artışları"
    g8.add_data(Reference(ws, min_col=6, min_row=47, max_row=52), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=5, min_row=48, max_row=52))
    g8.height, g8.width = 9, 12
    ws.add_chart(g8, "F54")
    baski_hazirla(ws, "A1:M52", URUN_AD)
    alt_bant(ws, 52, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 40, "B": 20, "C": 20, "D": 18, "E": 20, "F": 20, "G": 20, "H": 20, "I": 18})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Zam çarpanına göre senaryo bant aralığı ve hammadde duyarlılık (tornado) analizi.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Zam Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Toplam Maliyet Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        (7, "İyimser", "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiSenIyiZam,0)",
         "=IFERROR(HESAP!B16*UreticiSenIyiZam/UreticiSenBazZam,0)"),
        (8, "Baz", "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiSenBazZam,0)", "=HESAP!B16"),
        (9, "Kötümser", "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiSenKotuZam,0)",
         "=IFERROR(HESAP!B16*UreticiSenKotuZam/UreticiSenBazZam,0)"),
        (10, "Kritik", "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiSenKritikZam,0)",
         "=IFERROR(HESAP!B16*UreticiSenKritikZam/UreticiSenBazZam,0)"),
    ]
    for satir, ad, form, c in senaryolar:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 3, c, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Tornado Analizi", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 13, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 13, 2, "Düşük Senaryo (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 13, 3, "Yüksek Senaryo (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 13, 4, "Etki Genişliği (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    tornado = [
        (14, "Un Fiyatı", '=SUMIFS(tblRecete[Hammadde Tutarı (TL)],tblRecete[Hammadde],"Un")*UreticiTornadoUnOran',
         '=SUMIFS(tblRecete[Hammadde Tutarı (TL)],tblRecete[Hammadde],"Un")*UreticiTornadoUnOran*3'),
        (15, "Fire Oranı", '=SUM(E6:E9)*UreticiTornadoFireOran', '=SUM(E6:E9)*UreticiTornadoFireOran*3'),
    ]
    for satir, ad, dus, yuk in tornado:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, dus, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 3, yuk, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, f'=C{satir}-B{satir}', sayi=TL, yazi=DIS_REF_YEŞIL)
    baski_hazirla(ws, "A1:D10", URUN_AD)
    alt_bant(ws, 18, "Tornado genişliği hangi değişkenin sonucu en çok etkilediğini gösterir.")
    genislik(ws, {"A": 26, "B": 24, "C": 24, "D": 22})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Raporu", kalin=True, boyut=15, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "=CONCATENATE(\"Rapor Tarihi: \",TEXT(UreticiRaporTarihi,\"gg.aa.yyyy\"),\" | \",UreticiFirmaUnvan)",
      yazi=GRİ)
    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=13, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 7, 2, "=KARAR!B12", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[7].height = 40
    h(ws, 9, 1, "Ana Göstergeler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    rapor_satirlari = [
        (10, "Toplam Hammadde Maliyeti (TL)", "=HESAP!B13", TL),
        (11, "Toplam Zam Sonrası Maliyet (TL)", "=HESAP!B15", TL),
        (12, "Toplam Maliyet Artışı (TL)", "=HESAP!B16", TL),
        (13, "Ortalama Maliyet Artış Oranı", "=HESAP!B17", YÜZDE),
        (14, "Zam Yansıtma Gereksinimi", "=HESAP!B21", NORMAL),
        (15, "En Çok Etkilenen Ürün", "=HESAP!B22", NORMAL),
        (16, "Hedef Kâr Marjı", "=IF(COUNTA(tblRecete[Ürün Adı])>0,UreticiHedefMarjOran,0)", YÜZDE),
    ]
    for satir, ad, form, sayi in rapor_satirlari:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Varsayımlar ve Uyarılar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "Hedef kâr marjı, işçilik ve GÜG oranları AYARLAR'dan okunur; güncellenebilir.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 20, 1, "Fiyat kararı yönetim onayına tabidir; bu dosya karar destek aracıdır.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A20:P20")
    h(ws, 22, 1, "=CONCATENATE(\"Sürüm \",UreticiDosyaSurumu,\" | \",UreticiRaporHazirlayan)",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A22:P22")
    baski_hazirla(ws, "A1:P22", URUN_AD)
    genislik(ws, {"A": 40, "B": 26})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek Reçete Verileri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Ürünü tanımak için bu sayfadaki örnek kalemler kullanılabilir; RECETE_LISTESI "
                "üzerinde denemeler yapın.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ürün Kodu", "Ürün Adı", "Hammadde", "Birim", "Miktar",
                "Birim Fiyat (TL)", "Fire Oranı", "Zam Oranı",
                "Hammadde Tutarı (TL)", "Fire Dahil Tutar (TL)",
                "Zam Sonrası Birim Fiyat (TL)", "Zam Sonrası Tutar (TL)",
                "Fark (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Hammadde Tutarı (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="",0,tblOrnekRecete[[#This Row],[Miktar]]*tblOrnekRecete[[#This Row],[Birim Fiyat (TL)]])',
        "Fire Dahil Tutar (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="",0,tblOrnekRecete[[#This Row],[Miktar]]*(1+tblOrnekRecete[[#This Row],[Fire Oranı]])*tblOrnekRecete[[#This Row],[Birim Fiyat (TL)]])',
        "Zam Sonrası Birim Fiyat (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="",0,tblOrnekRecete[[#This Row],[Birim Fiyat (TL)]]*(1+tblOrnekRecete[[#This Row],[Zam Oranı]]))',
        "Zam Sonrası Tutar (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="",0,tblOrnekRecete[[#This Row],[Miktar]]*(1+tblOrnekRecete[[#This Row],[Fire Oranı]])*tblOrnekRecete[[#This Row],[Zam Sonrası Birim Fiyat (TL)]])',
        "Fark (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="",0,tblOrnekRecete[[#This Row],[Zam Sonrası Tutar (TL)]]-tblOrnekRecete[[#This Row],[Fire Dahil Tutar (TL)]])',
        "Açıklama": '=IF(tblOrnekRecete[[#This Row],[Ürün Adı]]="","",CONCATENATE("Örnek reçete: ",tblOrnekRecete[[#This Row],[Ürün Adı]]))',
    }
    tablo_ekle(ws, "tblOrnekRecete", "A5:N1004", sutunlar, formuller)
    ornek = [
        ("UR-01", "Sade Poğaça", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-01", "Sade Poğaça", "Margarin", "kg", 0.020, 90, 0.03, 0.08),
        ("UR-01", "Sade Poğaça", "Yumurta", "kg", 0.020, 45, 0.05, 0.05),
        ("UR-02", "Tam Buğday Ekmeği", "Un", "kg", 0.060, 20, 0.05, 0.15),
        ("UR-02", "Tam Buğday Ekmeği", "Tam Buğday Unu", "kg", 0.020, 25, 0.05, 0.12),
        ("UR-03", "Açma", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-03", "Açma", "Margarin", "kg", 0.030, 90, 0.03, 0.08),
        ("UR-04", "Simit", "Un", "kg", 0.050, 20, 0.05, 0.15),
        ("UR-04", "Simit", "Susam", "kg", 0.020, 60, 0.06, 0.06),
    ]
    for i, (kod, urun, ham, birim, miktar, fiyat, fire, zam) in enumerate(ornek, 6):
        for kolon, deger in [(1, kod), (2, urun), (3, ham), (4, birim), (5, miktar),
                             (6, fiyat), (7, fire), (8, zam)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(6, 1004):
        ws.cell(row=satir, column=5).number_format = CATI
        for kolon in (6, 9, 10, 11, 12, 13):
            ws.cell(row=satir, column=kolon).number_format = TL
        for kolon in (7, 8):
            ws.cell(row=satir, column=kolon).number_format = YÜZDE
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8])
    sabitle(ws, "A6")
    genislik(ws, {"A": 10, "B": 20, "C": 18, "D": 8, "E": 9, "F": 13, "G": 11,
                  "H": 11, "I": 16, "J": 18, "K": 20, "L": 18, "M": 13, "N": 52})


def degisiklik_kaydi(ws):
    sayfa_hazirla(ws, "DEGISIKLIK_KAYDI", "8D8D8D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Değişiklik Kaydı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 4, ["Tarih", "Sürüm", "Değişiklik", "Neden", "Yapan"])
    kayit = ["10.08.2026", "1.0.0", "İlk üretim", "Ürünün yayına hazırlanması", "ExcelArşiv"]
    for i, deger in enumerate(kayit, 1):
        h(ws, 5, i, deger, yazi="333333")
    genislik(ws, {"A": 16, "B": 12, "C": 40, "D": 32, "E": 18})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "808080", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Bağlı Listeler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Bu sayfadaki listeler açılır menüleri besler; yeni seçenek eklemek için "
                "sarı hücrelere yazın.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Hammaddeler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ham in enumerate(HAMMADDELER, 7):
        h(ws, i, 1, ham, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Hammadde | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni hammadde için satır ekleyin")
    h(ws, 7, 3, "Birimler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, birim in enumerate(BIRIMLER, 7):
        h(ws, i, 3, birim, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"C{i}", "Tanım: Birim | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni birim için satır ekleyin")
    genislik(ws, {"A": 22, "C": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Eşikler, senaryo çarpanları ve oranları "
                "güncelleyin; değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"])
    TARIH_D = date(2026, 8, 10)
    YURURLUK = date(2026, 8, 10)
    DOGRULAMA = date(2026, 8, 10)
    satirlar = [
        (7, "UreticiFirmaUnvan", "Örnek Firma A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "UreticiRaporTarihi", TARIH_D, "Tarih", None, None, "Raporun dayanak tarihi.", "Kullanıcı"),
        (9, "UreticiRaporHazirlayan", "Üretim ve Mali İşler Müdürü", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "UreticiDosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "UreticiHedefMarjOran", 0.30, "Oran", 0, 1, "Hedef kâr marjı; önerilen fiyat bu marjı koruyacak şekilde hesaplanır.", "Varsayım"),
        (12, "UreticiZamYansitmaEsik", 0.03, "Oran", 0, 1, "Üzerindeki genel maliyet artışı ZAM YANSIT kararı üretir.", "Varsayım"),
        (13, "UreticiFireKritikOran", 0.10, "Oran", 0, 1, "Üzerindeki fire oranı kritik kabul edilir.", "Varsayım"),
        (14, "UreticiIscilikOran", 0.12, "Oran", 0, 1, "Hammadde maliyetine uygulanan işçilik oranı.", "Varsayım"),
        (15, "UreticiGugOran", 0.08, "Oran", 0, 1, "Hammadde maliyetine uygulanan genel üretim gideri oranı.", "Varsayım"),
        (16, "UreticiSenIyiZam", 0.00, "Oran", -0.5, 0.5, "İyimser senaryo zam çarpanı.", "Varsayım"),
        (17, "UreticiSenBazZam", 0.03, "Oran", -0.5, 0.5, "Baz senaryo zam çarpanı.", "Varsayım"),
        (18, "UreticiSenKotuZam", 0.08, "Oran", -0.5, 0.5, "Kötümser senaryo zam çarpanı.", "Varsayım"),
        (19, "UreticiSenKritikZam", 0.15, "Oran", -0.5, 0.5, "Kritik senaryo zam çarpanı.", "Varsayım"),
        (20, "UreticiTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (21, "UreticiKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (22, "UreticiKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (23, "UreticiMalzemePayEsik", 0.50, "Oran", 0, 1, "Tek hammaddenin toplam maliyet içindeki payı için yoğunlaşma eşiği.", "Varsayım"),
        (24, "UreticiP90Oran", 0.90, "Oran", 0, 1, "Birim fiyat dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (25, "UreticiTemelKolonSayisi", 8, "Adet", 1, 20, "Veri kalitesi skorunda temel kabul edilen zorunlu giriş kolonu sayısı.", "Varsayım"),
        (26, "UreticiTornadoUnOran", 0.10, "Oran", 0, 0.5, "Un fiyatı tornado adımı (%10).", "Varsayım"),
        (27, "UreticiTornadoFireOran", 0.05, "Oran", 0, 0.5, "Fire oranı tornado adımı.", "Varsayım"),
        (28, "UreticiYuksekMaliyetEsik", 50000, "TL", 0, 1000000000, "Yüksek maliyet artışı kontrol eşiği.", "Varsayım"),
        (29, "UreticiAksiyon1_ACIKLAMA",
         "Genel maliyet artışı eşiğin üzerinde; en çok etkilenen ürünün fiyatını hedef marjı koruyacak şekilde güncelleyin.", "Metin", None, None,
         "ZAM YANSIT kararında gösterilen ilk öneri.", "Varsayım"),
        (30, "UreticiAksiyon2_ACIKLAMA",
         "Yüksek paylı hammaddenin tedarikini toplu alım veya alternatif tedarikçiyle müzakere ederek zam etkisini azaltın.", "Metin", None, None,
         "KISMİ YANSIT kararında gösterilen ikinci öneri.", "Varsayım"),
        (31, "UreticiAksiyon3_ACIKLAMA",
         "Fire oranı yüksek hammaddelerde üretim sürecini gözden geçirin; fire azaltımı maliyeti doğrudan iyileştirir.", "Metin", None, None,
         "FİYATI KORU kararında gösterilen üçüncü öneri.", "Varsayım"),
    ]
    for satir, anahtar, deger, birim, alt, ust, aciklama, kaynak in satirlar:
        h(ws, satir, 1, anahtar, kalin=True, zemin=ACIK_GRI, boyut=9, kaydir=True)
        hb = h(ws, satir, 2, deger, yazi=GIRIS_YAZI, zemin=GIRIS_SARI, kalin=True)
        h(ws, satir, 3, birim, boyut=9, yazi=GRİ)
        h(ws, satir, 4, alt, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 5, ust, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 6, aciklama, boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, kaynak, boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        if isinstance(deger, date):
            hb.number_format = TARİH
        elif isinstance(deger, (int, float)):
            hb.number_format = TL if birim == "TL" else YÜZDE if birim == "Oran" else CATI
        if isinstance(deger, (int, float)) and alt is not None:
            dv = DataValidation(type="decimal", operator="between", formula1=str(alt),
                                formula2=str(ust), allow_blank=False,
                                promptTitle=anahtar, prompt=aciklama,
                                showErrorMessage=True,
                                errorTitle="Değer aralık dışı", error=aciklama)
            ws.add_data_validation(dv)
            dv.add(f"B{satir}")
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: {aciklama} | Neden önemli: Tüm hesaplamaları etkiler | "
                   f"Doğru kullanım: Açıklamaya uygun değer girin | Kaynak: {kaynak}")
        if 29 <= satir <= 31:
            ws.cell(row=satir, column=2).alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[satir].height = 48
        ws.row_dimensions[satir].height = 26
    for i in range(1, 13):
        satir = 32 + i
        h(ws, satir, 1, f"UreticiPanelSatir_{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"HESAP/PANO'daki grafiğin ürün sırası ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel sıra numarası {i} | "
                   f"Neden önemli: Grafikleri besler | "
                   f"Doğru kullanım: Ürün sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 34, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 22, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "RECETE_LISTESI sayfasında sarı hücrelere reçete kalemlerinizi girin. Ürün kodu/adı, "
         "hammadde, birim, miktar, birim fiyat, fire ve zam oranları doldurulunca tutarlar "
         "otomatik hesaplanır."),
        ("Zam yansıtma mantığı",
         "Hammadde zam oranı, birim fiyatın üzerine uygulanır ve zam sonrası tutarı üretir. "
         "Fark, zam sonrası tutar ile fire dahil eski tutar arasındaki farktır; ürün bazında "
         "toplanarak genel maliyet artışına dönüşür."),
        ("Önerilen satış fiyatı",
         "Hedef kâr marjı AYARLAR'dan okunur; önerilen fiyat, işçilik ve GÜG dahil zam sonrası "
         "birim maliyetin marja göre fiyatlanmasıyla bulunur (maliyet / (1 - marj))."),
        ("Karar ve aksiyon",
         "KARAR sayfası maliyet artışını eşikle karşılaştırır; ZAM YANSIT / KISMİ YANSIT / "
         "FİYATI KORU kararını gerekçesiyle üretir ve önerilen aksiyonları gösterir."),
        ("Oranlar ve senaryolar",
         "AYARLAR sayfasından hedef marj, işçilik, GÜG, senaryo ve tornado çarpanlarını "
         "güncelleyin; PANO ve SENARYO_DUYARLILIK anında güncellenir."),
        ("Güven ve karar kalitesi",
         "Veri kalite skoru girişlerin tamlığını ölçer. Reçete kalemi yokken karar VERİ YOK; "
         "negatif girişte DURDUR üretir. Verileriniz cihazınızdan çıkmaz; dosya tamamen çevrimdışıdır."),
        ("Uyumluluk",
         "Excel 2016-365 (Windows ve Mac), LibreOffice ve Google Sheets ile uyumludur. "
         "Makro içermez; formüller ayrıştırıcı ile denetlenmiştir."),
        ("Karar kuralları",
         "KARAR sayfasındaki sonuç şu kurallarla üretilir: (1) reçete kalemi yoksa VERİ YOK; "
         "(2) negatif birim fiyat veya miktar varsa DURDUR; (3) genel maliyet artışı eşiğin "
         "üzerindeyse ZAM YANSIT; (4) artış pozitif ancak eşiğin altındaysa KISMİ YANSIT; "
         "(5) aksi durumda FİYATI KORU. Gerekçe, kararın altında makine tarafından üretilir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        ws.row_dimensions[i].height = 56
    genislik(ws, {"A": 34, "B": 110})


def kosullu_bicimlendirme(wb):
    kirmizi = PatternFill("solid", start_color="FDE9E9", end_color="FDE9E9")
    yesil = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
    sari = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
    kirmizi_font = Font(name=FONT, color="B3261E", bold=True, size=10)
    yesil_font = Font(name=FONT, color="1F7A4D", bold=True, size=10)

    ws = wb["KARAR"]
    for deger, dolu, yazifont in (("VERİ YOK", yesil, yesil_font), ("DURDUR", kirmizi, kirmizi_font),
                                  ("ZAM YANSIT", kirmizi, kirmizi_font), ("KISMİ YANSIT", sari, None),
                                  ("FİYATI KORU", yesil, yesil_font)):
        ws.conditional_formatting.add("B11",
            CellIsRule(operator="equal", formula=[f'"{deger}"'], fill=dolu, font=yazifont))

    ws = wb["PANO"]
    for r in (4, 6):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"ZAM YANSIT"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"KISMİ YANSIT"'], fill=sari))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"FİYATI KORU"'], fill=yesil, font=yesil_font))
    for c in ("B", "C", "D", "E", "F", "G", "H", "I"):
        for r in range(11, 15):
            ws.conditional_formatting.add(f"{c}{r}",
                CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

    ws = wb["RECETE_LISTESI"]
    for c in ("E", "F"):
        ws.conditional_formatting.add(f"{c}6:{c}1004",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("G6:G1004",
        CellIsRule(operator="greaterThan", formula=["UreticiFireKritikOran"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("H6:H1004",
        CellIsRule(operator="greaterThan", formula=["0.5"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("M6:M1004",
        CellIsRule(operator="greaterThan", formula=["0"], fill=sari, font=kirmizi_font))

    ws = wb["HESAP"]
    for hucre in ("B16",):
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="greaterThan", formula=["UreticiYuksekMaliyetEsik"],
                       fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("H6:H9",
        CellIsRule(operator="greaterThan", formula=["UreticiZamYansitmaEsik"], fill=sari, font=kirmizi_font))

    ws = wb["KONTROLLER"]
    for r in range(6, 12):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="notEqual", formula=['"TEMİZ"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yesil_font))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="between", formula=["70", "89"], fill=sari))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kirmizi_font))


def moduller(wb):
    # Analitik modül değer/yorumları KONTROLLER sayfasında hücre olarak durur;
    # ad tanımları bu hücrelere işaret eder (D03/D04/D05 gereği).
    ws = wb["KONTROLLER"]
    modul_satirlari = {
        "modulT1OrtMarj": ("T1", "Ortalama Fire Oranı",
                           "=IFERROR(AVERAGE(tblRecete[Fire Oranı]),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama fire oranı ",TEXT(C24,"0.0%")," — fire maliyet yönetiminin genel düzeyini gösterir")',
                           YÜZDE),
        "modulT2OrtZam": ("T2", "Ortalama Zam Oranı",
                          "=IFERROR(AVERAGE(tblRecete[Zam Oranı]),0)",
                          '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama hammadde zam oranı ",TEXT(C25,"0.0%")," — genel maliyet artışının ana kaynağını gösterir")',
                          YÜZDE),
        "modulO1KarSapmasi": ("O1", "Kalem Fark Sapması (MAD)",
                              "=IFERROR(AVEDEV(tblRecete[Fark (TL)]),0)",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Kalem bazında maliyet farkı sapması ",TEXT(C26,"₺ #,##0")," TL — zam etkisinin kalemler arası heterojenliğini ölçer")',
                              TL),
        "modulO2EnYuksekFark": ("O2", "En Yüksek Kalem Farkı",
                                "=IFERROR(MAX(tblRecete[Fark (TL)]),0)",
                                '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek kalem farkı ",TEXT(C27,"₺ #,##0")," TL — bu hammadde zam yansıtmasının öncelikli kalemidir")',
                                TL),
        "modulO3Senaryo": ("O3", "Senaryo Bant Genişliği",
                           "=IFERROR(STDEV.P(SENARYO_DUYARLILIK!C7:C10),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo maliyet artışı standart sapması ",TEXT(C28,"₺ #,##0")," TL — zam belirsizliğinin etkisini gösterir")',
                           TL),
        "modulO6Hhi": ("O6", "Maliyet Yoğunlaşması (HHI)",
                       '=IFERROR(IF(SUM(HESAP!$D$6:$D$9)=0,0,SUMPRODUCT((HESAP!$D$6:$D$9/SUM(HESAP!$D$6:$D$9))^2)),0)',
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Ürün bazlı maliyet yoğunlaşma endeksi ",TEXT(C29,"0.000")," — 1 e yaklaştıkça maliyet tek üründe toplanır")',
                       None),
        "modulO8KaliteSkor": ("O8", "Veri Kalite Skoru",
                              "=KONTROLLER!B16",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C30,"0")," puan — ",IF(C30>=UreticiKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))',
                              CATI),
        "modulI1Tahmin": ("I1", "Gelecek Dönem Artış Tahmini",
                          "=HESAP!B39",
                          '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem maliyet artışı tahmini ",TEXT(C31,"₺ #,##0")," TL — bant aralığı için alt/üst sınır HESAP sayfasındadır")',
                          TL),
        "modulI2P90": ("I2", "Birim Fiyat Yüzdelik P90",
                       "=IFERROR(_xlfn.PERCENTILE.INC(tblRecete[Birim Fiyat (TL)],UreticiP90Oran),0)",
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Birim fiyatların %90 ı ",TEXT(C32,"₺ #,##0")," TL altındadır — fiyat dağılımının üst sınırını gösterir")',
                       TL),
    }
    baslik_satiri(ws, 23, ["Modül", "Gösterge", "Değer", "Makine Yorumu"])
    for i, (ad, (kod, gosterge, form, yorum, sayi)) in enumerate(modul_satirlari.items()):
        satir = 24 + i
        h(ws, satir, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, satir, 2, gosterge, yazi="333333")
        h(ws, satir, 3, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[satir].height = 30
        ad_ekle(wb, ad, f"KONTROLLER!$C${satir}")
        ad_ekle(wb, ad + "Yorum", f"KONTROLLER!$D${satir}")


def main():
    wb = Workbook()
    sayfa_sirasi = ["KAPAK", "HIZLI_BASLANGIC", "RECETE_LISTESI", "HESAP", "KONTROLLER",
                    "KARAR", "PANO", "SENARYO_DUYARLILIK", "RAPOR", "ORNEK_VERI",
                    "DEGISIKLIK_KAYDI", "LISTELER", "AYARLAR", "KILAVUZ"]
    for ad in sayfa_sirasi:
        wb.create_sheet(ad)
    ilk = wb["Sheet"]
    wb.remove(ilk)

    kapak(wb["KAPAK"])
    hizli_baslangic(wb["HIZLI_BASLANGIC"])
    recete_listesi(wb["RECETE_LISTESI"])
    hesap(wb["HESAP"])
    kontroller(wb["KONTROLLER"])
    karar(wb["KARAR"])
    pano(wb["PANO"])
    senaryo_duyarlilik(wb["SENARYO_DUYARLILIK"])
    rapor(wb["RAPOR"])
    ornek_veri(wb["ORNEK_VERI"])
    degisiklik_kaydi(wb["DEGISIKLIK_KAYDI"])
    listeler(wb["LISTELER"])
    ayarlar(wb["AYARLAR"])
    kilavuz(wb["KILAVUZ"])

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=60)
    moduller(wb)
    kosullu_bicimlendirme(wb)

    # AYARLAR anahtar değerleri için ad tanımları (D10/D13, G07)
    for satir in range(7, 45):
        anahtar = wb["AYARLAR"].cell(row=satir, column=1).value
        if isinstance(anahtar, str) and anahtar.startswith("Uretici"):
            ad_ekle(wb, anahtar, f"AYARLAR!$B${satir}")

    # Kritik çıktı ad tanımları (D02 karşılıkları ve PANO beslemeleri)
    ad_ekle(wb, "UretimToplamZamSonrasiMaliyet", "HESAP!$B$15")
    ad_ekle(wb, "UretimEnYuksekArtis", "HESAP!$B$23")
    ad_ekle(wb, "UretimKarar", "KARAR!$B$11")
    ad_ekle(wb, "UretimGenelArtisOrani", "HESAP!$B$20")
    ad_ekle(wb, "UretimHedefFiyatEnYuksek", "HESAP!$I$9")
    ad_ekle(wb, "UretimProjX", "HESAP!$A$46:$A$49")

    ad_ekle(wb, "ListeHammaddeler", "LISTELER!$A$7:$A$15")
    ad_ekle(wb, "ListeBirimler", "LISTELER!$C$7:$C$12")

    for ws in wb.worksheets:
        sayfa_koru(ws)

    dosya = "UretimRecetesiVeZamYansitmaHesaplayici.xlsx"
    wb.save(dosya)
    print("Dosya oluşturuldu:", dosya)


if __name__ == "__main__":
    main()
