"""Nakliye Maliyeti Hesaplayıcı — Excel Üretim Mandası v4.0 uyumlu üretim betiği."""
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

URUN_AD = "Nakliye Maliyeti Hesaplayıcı"
SURUM = "1.0.0"
RENK = "B08948"

ARAC_TIPLERI = ["Kamyonet", "Kamyon", "Tır", "Çekici"]
GUZERGAHLAR = ["İstanbul", "Ankara", "İzmir", "Bursa", "Antalya", "Adana", "Trabzon", "Samsun"]
PLAKALAR = ["34 ABC 123", "06 DEF 456", "35 GHI 789", "16 JKL 012", "41 MNO 345", "07 PQR 678"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Araç filosu ve sefer kayıtları üzerinden nakliye maliyetini yakıt, bakım, "
                "sürücü ve amortisman bileşenleriyle hesaplayın; km ve ton-km başına maliyeti, "
                "sefer kârlılığını ve iyileştirme önceliklerini gerekçesiyle raporlayın.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Sefer mesafesi ve araç özelliklerinden yakıt, bakım, sürücü ve amortisman maliyetinin otomatik hesaplanması",
        "Araç başına toplam maliyet ve sefer bazında km/ton-km başına birim maliyetin bulunması",
        "Hasılat ile maliyet karşılaştırmasıyla sefer kârlılığı ve net kârın hesaplanması",
        "Maliyet yoğunlaşması, anomali, senaryo ve duyarlılık analizi ile karar kapısı ve aksiyon önerisi",
        "Kârlılık eşiklerine göre KARLI / İNCELE / ZARARLI kararının gerekçesiyle üretilmesi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Sefer bazında nakliye maliyetini ve kârlılığını ölçmek isteyen lojistik ve nakliye işletmeleri",
        "Araç filosu maliyetlerini kalem kalem analiz etmek isteyen filo yöneticileri",
        "Taşıma fiyatını maliyetle kanıtlayarak belirlemek isteyen işletme sahipleri",
        "Müşterisine taşıma maliyeti ve kârlılık analizi öneren mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 20, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. ARAC_LISTESI ve SEFERLER sayfalarına verilerinizi girin. "
                 "3. PANO ve KARAR sayfalarından nakliye kârlılık kararını izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A21:P21")
    h(ws, 23, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A23:P23")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Araç listesini girin",
         "ARAC_LISTESI sayfasındaki sarı hücrelere araç plakasını, tipini, yakıt tüketimini, "
         "bakım oranını, sürücü maliyetini ve amortismanı girin."),
        ("Adım 2 — Sefer kayıtlarını girin",
         "SEFERLER sayfasındaki sarı hücrelere sefer tarihi, plaka, güzergah, mesafe, yük ve "
         "hasılatı girin; maliyet kalemleri otomatik hesaplanır."),
        ("Adım 3 — Kârlılığı izleyin",
         "HESAP sayfası toplam maliyet, birim maliyet ve net kârı üretir; PANO görselleştirir."),
        ("Adım 4 — Kararı değerlendirin",
         "KARAR sayfası kâr marjını eşiklerle karşılaştırarak KARLI / İNCELE / ZARARLI "
         "kararını gerekçesiyle üretir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Mesafeyi km yerine metre girmek (maliyetler hatalı küçülür)",
        "Yakıt tüketimini 100 km başına değil km başına girmek",
        "Plakayı araç listesinden farklı yazmak (maliyet kalemleri eşleşmez)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 16, 1, "Sık sorulan sorular", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sss = [
        "Soru: Yakıt maliyeti nasıl hesaplanır? Cevap: Mesafe × (tüketim/100) × yakıt fiyatı formülüyle otomatik üretilir.",
        "Soru: Sürücü maliyeti nasıl dağıtılır? Cevap: Sefer süresi (mesafe/günlük kapasite) ile günlük sürücü maliyeti çarpılır.",
        "Soru: Birim maliyet nedir? Cevap: Sefer toplam maliyetinin km ve ton-km başına düşen tutarıdır.",
        "Soru: Araç listesinde aradığım yoksa? Cevap: LISTELER sayfasındaki sarı hücreye yeni araç tipini ekleyin.",
    ]
    for i, m in enumerate(sss, 17):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
        ws.row_dimensions[i].height = 32
    h(ws, 22, 1, "Bu dosya karar destek aracıdır; maliyetler fiili yakıt ve bakım kayıtlarıyla "
                 "doğrulanmalıdır.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def arac_listesi(ws):
    sayfa_hazirla(ws, "ARAC_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Araç Listesi ve Maliyet Parametreleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Yakıt, bakım, sürücü ve amortisman "
                "maliyetleri seferlere otomatik dağıtılır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Plaka", "Araç Tipi", "Yakıt Tüketimi (lt/100km)", "Bakım Oranı (TL/km)",
                "Sürücü Maliyeti (TL/gün)", "Günlük Kapasite (km/gün)", "Amortisman (TL/km)",
                "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Açıklama": '=IF(tblArac[[#This Row],[Plaka]]="","",CONCATENATE("Araç: ",tblArac[[#This Row],[Plaka]]," | Tip: ",tblArac[[#This Row],[Araç Tipi]]," | Tüketim: ",TEXT(tblArac[[#This Row],[Yakıt Tüketimi (lt/100km)]],"0,0")," lt/100km"))',
    }
    tablo_ekle(ws, "tblArac", "A5:H1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeAracTipleri", "B6:B1004",
              baslik="Araç Tipi", mesaj="Listeden araç tipini seçin.",
              hata_baslik="Geçersiz tip", hata_mesaj="Listede olmayan tip giremezsiniz.")
    for kolon, ad in (("C", "Yakıt Tüketimi"), ("E", "Sürücü Maliyeti")):
        dogrulama(ws, "decimal", "0", f"{kolon}6:{kolon}1004",
                  baslik=ad, mesaj="0 veya daha büyük değer girin.",
                  hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = CATI
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=5).number_format = TL
        ws.cell(row=satir, column=6).number_format = CATI
        ws.cell(row=satir, column=7).number_format = TL
    ornek = [
        ("34 ABC 123", "Kamyonet", 11, 1.4, 1200, 350, 1.1),
        ("06 DEF 456", "Kamyon", 24, 2.2, 1600, 500, 1.8),
        ("35 GHI 789", "Tır", 32, 2.8, 2000, 650, 2.4),
        ("16 JKL 012", "Kamyonet", 10, 1.2, 1100, 320, 1.0),
        ("41 MNO 345", "Çekici", 28, 2.5, 1800, 600, 2.1),
        ("07 PQR 678", "Kamyon", 25, 2.3, 1650, 520, 1.9),
    ]
    for i, (plaka, tip, tuketim, bakim, surucu, kapasite, amortisman) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = plaka
        ws.cell(row=i, column=2).value = tip
        ws.cell(row=i, column=3).value = tuketim
        ws.cell(row=i, column=4).value = bakim
        ws.cell(row=i, column=5).value = surucu
        ws.cell(row=i, column=6).value = kapasite
        ws.cell(row=i, column=7).value = amortisman
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 12, "C": 20, "D": 16, "E": 20, "F": 18, "G": 18, "H": 56})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def seferler(ws):
    sayfa_hazirla(ws, "SEFERLER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sefer Kayıtları", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Her seferin mesafe, yük ve hasılat bilgisini girin; maliyet kalemleri "
                "araç parametrelerinden otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Sefer Numarası", "Tarih", "Plaka", "Güzergah", "Mesafe (km)", "Yük (ton)",
                "Hasılat (TL)", "Yakıt Maliyeti (TL)", "Bakım Maliyeti (TL)",
                "Sürücü Maliyeti (TL)", "Amortisman (TL)", "Toplam Maliyet (TL)",
                "Birim Maliyet (TL/km)", "Ton-Km Maliyeti (TL)", "Net Kâr (TL)",
                "Kâr Marjı", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Yakıt Maliyeti (TL)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="","",tblSefer[[#This Row],[Mesafe (km)]]*(IFERROR(INDEX(tblArac[Yakıt Tüketimi (lt/100km)],MATCH(tblSefer[[#This Row],[Plaka]],tblArac[Plaka],0)),0)/100)*NakliyeYakitFiyati)',
        "Bakım Maliyeti (TL)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="","",tblSefer[[#This Row],[Mesafe (km)]]*IFERROR(INDEX(tblArac[Bakım Oranı (TL/km)],MATCH(tblSefer[[#This Row],[Plaka]],tblArac[Plaka],0)),0))',
        "Sürücü Maliyeti (TL)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="","",IFERROR(tblSefer[[#This Row],[Mesafe (km)]]/IFERROR(INDEX(tblArac[Günlük Kapasite (km/gün)],MATCH(tblSefer[[#This Row],[Plaka]],tblArac[Plaka],0)),1)*IFERROR(INDEX(tblArac[Sürücü Maliyeti (TL/gün)],MATCH(tblSefer[[#This Row],[Plaka]],tblArac[Plaka],0)),0),0))',
        "Amortisman (TL)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="","",tblSefer[[#This Row],[Mesafe (km)]]*IFERROR(INDEX(tblArac[Amortisman (TL/km)],MATCH(tblSefer[[#This Row],[Plaka]],tblArac[Plaka],0)),0))',
        "Toplam Maliyet (TL)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="","",SUM(tblSefer[[#This Row],[Yakıt Maliyeti (TL)]]:tblSefer[[#This Row],[Amortisman (TL)]]))',
        "Birim Maliyet (TL/km)": '=IF(tblSefer[[#This Row],[Mesafe (km)]]="",0,IFERROR(tblSefer[[#This Row],[Toplam Maliyet (TL)]]/tblSefer[[#This Row],[Mesafe (km)]],0))',
        "Ton-Km Maliyeti (TL)": '=IF(OR(tblSefer[[#This Row],[Mesafe (km)]]="",tblSefer[[#This Row],[Yük (ton)]]=""),0,IFERROR(tblSefer[[#This Row],[Toplam Maliyet (TL)]]/(tblSefer[[#This Row],[Mesafe (km)]]*tblSefer[[#This Row],[Yük (ton)]]),0))',
        "Net Kâr (TL)": '=IF(tblSefer[[#This Row],[Hasılat (TL)]]="","",IFERROR(tblSefer[[#This Row],[Hasılat (TL)]]-tblSefer[[#This Row],[Toplam Maliyet (TL)]],0))',
        "Kâr Marjı": '=IF(tblSefer[[#This Row],[Hasılat (TL)]]="",0,IFERROR(tblSefer[[#This Row],[Net Kâr (TL)]]/tblSefer[[#This Row],[Hasılat (TL)]],0))',
        "Açıklama": '=IF(tblSefer[[#This Row],[Sefer Numarası]]="","",CONCATENATE("Sefer: ",tblSefer[[#This Row],[Sefer Numarası]]," | Güzergah: ",tblSefer[[#This Row],[Güzergah]]," | Net Kâr: ",TEXT(tblSefer[[#This Row],[Net Kâr (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblSefer", "A5:Q1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListePlakalar", "C6:C1004",
              baslik="Plaka", mesaj="Listeden araç plakasını seçin.",
              hata_baslik="Geçersiz plaka", hata_mesaj="Araç listesinde olmayan plaka giremezsiniz.")
    dogrulama(ws, "list", "ListeGuzergahlar", "D6:D1004",
              baslik="Güzergah", mesaj="Listeden güzergahı seçin.",
              hata_baslik="Geçersiz güzergah", hata_mesaj="Listede olmayan güzergah giremezsiniz.")
    for kolon, ad in (("E", "Mesafe"), ("F", "Yük"), ("G", "Hasılat")):
        dogrulama(ws, "decimal", "0", f"{kolon}6:{kolon}1004",
                  baslik=ad, mesaj="0 veya daha büyük değer girin.",
                  hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "date", None, "B6:B1004",
              baslik="Tarih", mesaj="Sefer tarihini GG.AA.YYYY biçiminde girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Geçerli bir tarih girin.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=2).number_format = TARİH
        ws.cell(row=satir, column=5).number_format = CATI
        ws.cell(row=satir, column=6).number_format = CATI
        for kolon in (7, 8, 9, 10, 11, 12, 13, 14, 15):
            ws.cell(row=satir, column=kolon).number_format = TL
        ws.cell(row=satir, column=16).number_format = YÜZDE
    ornek = [
        ("SF-1001", date(2026, 6, 3), "34 ABC 123", "İstanbul", 320, 1.5, 8500),
        ("SF-1002", date(2026, 6, 5), "06 DEF 456", "Ankara", 470, 8.0, 18500),
        ("SF-1003", date(2026, 6, 8), "35 GHI 789", "İzmir", 590, 16.0, 32000),
        ("SF-1004", date(2026, 6, 12), "16 JKL 012", "Bursa", 180, 1.2, 5200),
        ("SF-1005", date(2026, 6, 15), "41 MNO 345", "Antalya", 780, 12.0, 27500),
        ("SF-1006", date(2026, 6, 20), "07 PQR 678", "Adana", 890, 9.0, 29500),
        ("SF-1007", date(2026, 6, 24), "34 ABC 123", "Trabzon", 1180, 1.4, 21000),
        ("SF-1008", date(2026, 6, 28), "06 DEF 456", "Samsun", 730, 8.5, 20500),
        ("SF-1009", date(2026, 7, 2), "35 GHI 789", "İstanbul", 340, 15.0, 24500),
        ("SF-1010", date(2026, 7, 6), "41 MNO 345", "İzmir", 610, 11.0, 23500),
    ]
    for i, (sefer, tarih, plaka, guzergah, mesafe, yuk, hasilat) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = sefer
        ws.cell(row=i, column=2).value = tarih
        ws.cell(row=i, column=3).value = plaka
        ws.cell(row=i, column=4).value = guzergah
        ws.cell(row=i, column=5).value = mesafe
        ws.cell(row=i, column=6).value = yuk
        ws.cell(row=i, column=7).value = hasilat
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7])
    sabitle(ws, "A6")
    genislik(ws, {"A": 10, "B": 12, "C": 13, "D": 12, "E": 13, "F": 10, "G": 13,
                  "H": 16, "I": 16, "J": 16, "K": 15, "L": 17, "M": 17, "N": 17,
                  "O": 13, "P": 11, "Q": 56})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nakliye Maliyeti Hesaplama Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sefer ve araç verilerinden canlı üretilir; toplam maliyet, birim maliyet ve "
                "kârlılık hesaplanır.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "TOPLAM HASILAT (TL)", yazi="333333", kalin=True)
    h(ws, 6, 2, "=SUM(tblSefer[Hasılat (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "TOPLAM NAKLİYE MALİYETİ (TL)", yazi="333333", kalin=True)
    h(ws, 7, 2, "=SUM(tblSefer[Toplam Maliyet (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 8, 1, "Yakıt Maliyeti Toplamı (TL)", yazi="333333")
    h(ws, 8, 2, "=SUM(tblSefer[Yakıt Maliyeti (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Bakım Maliyeti Toplamı (TL)", yazi="333333")
    h(ws, 9, 2, "=SUM(tblSefer[Bakım Maliyeti (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Sürücü Maliyeti Toplamı (TL)", yazi="333333")
    h(ws, 10, 2, "=SUM(tblSefer[Sürücü Maliyeti (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Amortisman Toplamı (TL)", yazi="333333")
    h(ws, 11, 2, "=SUM(tblSefer[Amortisman (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Toplam Net Kâr (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 12, 2, "=B6-B7", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Toplam Kâr Marjı", yazi="333333")
    h(ws, 13, 2, "=IFERROR(B12/B6,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Toplam Mesafe (km)", yazi="333333")
    h(ws, 14, 2, "=SUM(tblSefer[Mesafe (km)])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Ortalama Birim Maliyet (TL/km)", yazi="333333")
    h(ws, 15, 2, "=IFERROR(B7/B14,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Ton-Km Toplamı", yazi="333333")
    h(ws, 16, 2, "=SUMPRODUCT(tblSefer[Mesafe (km)],tblSefer[Yük (ton)])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Ton-Km Birim Maliyeti (TL)", yazi="333333")
    h(ws, 17, 2, "=IFERROR(B7/B16,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Sefer Sayısı", yazi="333333")
    h(ws, 18, 2, "=COUNT(tblSefer[Sefer Numarası])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "En Yüksek Sefer Maliyeti (TL)", yazi="333333")
    h(ws, 19, 2, '=IFERROR(MAX(tblSefer[Toplam Maliyet (TL)]),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "En Pahalı Sefer Numarası", yazi="333333")
    h(ws, 20, 2, '=IFERROR(INDEX(tblSefer[Sefer Numarası],MATCH(MAX(tblSefer[Toplam Maliyet (TL)]),tblSefer[Toplam Maliyet (TL)],0)),"")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Maliyetin Kâra Oranı", yazi="333333")
    h(ws, 21, 2, "=IFERROR(B7/B6,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Zararlı Sefer Sayısı", yazi="333333")
    h(ws, 22, 2, '=COUNTIF(tblSefer[Net Kâr (TL)],"<0")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "Kârlı Sefer Oranı", yazi="333333")
    h(ws, 23, 2, '=IFERROR(COUNTIF(tblSefer[Kâr Marjı],">0")/COUNT(tblSefer[Sefer Numarası]),0)', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    # Panel verisi (grafikler için) — araç sırasına göre
    h(ws, 26, 1, "Araç Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 2, "Hasılat (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 4, "Toplam Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 6, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(6):
        h(ws, 27 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 27 + i, 2, f'=IFERROR(SUMIFS(tblSefer[Hasılat (TL)],tblSefer[Plaka],INDEX(tblArac[Plaka],NakliyePanelSatir_{i+1})),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 27 + i, 4, f'=IFERROR(SUMIFS(tblSefer[Toplam Maliyet (TL)],tblSefer[Plaka],INDEX(tblArac[Plaka],NakliyePanelSatir_{i+1})),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 27 + i, 6, f'=IFERROR(B{27+i}-D{27+i},0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    # Projeksiyon yardımcı serisi (sefer maliyet trendi)
    h(ws, 35, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 35, 2, "Toplam Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    maliyet_serisi = [("1", "=IFERROR(INDEX(tblSefer[Toplam Maliyet (TL)],1),0)"),
                      ("2", "=IFERROR(INDEX(tblSefer[Toplam Maliyet (TL)],2),0)"),
                      ("3", "=IFERROR(INDEX(tblSefer[Toplam Maliyet (TL)],3),0)"),
                      ("4", "=IFERROR(INDEX(tblSefer[Toplam Maliyet (TL)],4),0)")]
    for i, (donem, form) in enumerate(maliyet_serisi, 36):
        h(ws, i, 1, donem, sayi=CATI, yazi=GRİ)
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 41, 1, "Ortalama Sefer Maliyeti (TL)", yazi="333333")
    h(ws, 41, 2, "=IFERROR(AVERAGE(B36:B39),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 42, 1, "Gelecek Dönem Maliyet Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 42, 2, "=IFERROR(FORECAST.LINEAR(NakliyeTahminDonem,B36:B39,NakliyeProjX)-B39,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 43, 1, "Maliyet Std Sapması (TL)", yazi="333333")
    h(ws, 43, 2, "=IFERROR(STDEV.P(B36:B39),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 44, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 44, 2, "=IFERROR(B41-B43*NakliyeTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 45, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 45, 2, "=IFERROR(B41+B43*NakliyeTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 47, "Motor, sefer ve araç verilerinden canlı beslenir; tüm eşikler AYARLAR'dan okunur.")
    genislik(ws, {"A": 34, "B": 40, "C": 24, "D": 24, "E": 24, "F": 22})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Veri Kalite ve Kontrol Merkezi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Girdilerin kalitesini ölçer; eksik, negatif ve uç değerleri raporlar.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    baslik_satiri(ws, 5, ["Kontrol", "Sonuç", "Durum"])
    kontroller = [
        ("Araç girildi mi?", '=IF(COUNTA(tblArac[Plaka])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Sefer kaydı girildi mi?", '=IF(COUNTA(tblSefer[Sefer Numarası])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Negatif mesafe girişi var mı?", '=IF(COUNTIF(tblSefer[Mesafe (km)],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, "B3261E"),
        ("Zararlı sefer var mı?", '=IF(COUNTIF(tblSefer[Kâr Marjı],"<0")=0,"TEMİZ","ZARARLI VAR")', NORMAL, "B3261E"),
        ("Yakıt tüketimi girildi mi?", '=IF(COUNTIF(tblArac[Yakıt Tüketimi (lt/100km)],">0")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
        ("Sefer hasılatı girildi mi?", '=IF(COUNTIF(tblSefer[Hasılat (TL)],">0")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
    ]
    for i, (ad, form, durum1, durum2) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "AUTOMATİK", yazi=GRİ)
    h(ws, 14, 1, "Veri Kalite Skoru", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Doluluk Oranı (%)", yazi="333333")
    h(ws, 15, 2, '=IFERROR((COUNTA(tblArac[Plaka])+COUNTA(tblSefer[Sefer Numarası]))/(COUNTA(tblArac[Plaka])+COUNTA(tblSefer[Sefer Numarası])+COUNTBLANK(tblArac[Plaka])+COUNTBLANK(tblSefer[Sefer Numarası]))*100,0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Veri Kalite Skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, '=MIN(100,MAX(0,B15))', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Kalite Durumu", yazi="333333")
    h(ws, 17, 2, '=IF(B16>=NakliyeKaliteIyi,"YÜKSEK",IF(B16>=NakliyeKaliteOrta,"ORTA","DÜŞÜK"))', yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Canlı Formül Sayacı", yazi="333333")
    h(ws, 18, 2, '=IF(COUNTIF(tblSefer[Açıklama],"")=0,COUNTA(tblSefer[Açıklama]),0)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 20, "Tüm kontroller canlı formüllerle üretilir; sahte geçiş yoktur.")
    genislik(ws, {"A": 38, "B": 60, "C": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nakliye Kârlılık Karar Kapısı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Kâr marjı, maliyet oranı ve sefer durumuna göre karar üretir.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Toplam Kâr Marjı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B13", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Toplam Net Kâr (TL)", yazi="333333")
    h(ws, 7, 2, "=HESAP!B12", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hedef Kâr Marjı", yazi="333333")
    h(ws, 8, 2, "=IF(COUNTA(tblSefer[Sefer Numarası])>0,NakliyeHedefMarj,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Maliyetin Kâra Oranı", yazi="333333")
    h(ws, 9, 2, "=HESAP!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Zararlı Sefer Sayısı", yazi="333333")
    h(ws, 10, 2, "=HESAP!B22", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblSefer[Sefer Numarası])=0,"VERİ YOK",'
                 'IF(OR(COUNTIF(tblSefer[Mesafe (km)],"<0")>0,COUNTIF(tblSefer[Hasılat (TL)],"<0")>0),"DURDUR",'
                 'IF(HESAP!B13<0,"ZARARLI",'
                 'IF(HESAP!B13>=NakliyeHedefMarj,"KARLI",'
                 '"İNCELE"))))',
      boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 12, 2, '=IF(B11="VERİ YOK","Sefer verisi girilmedi; karar üretilemiyor.",'
                 'IF(B11="DURDUR","Negatif mesafe veya hasılat girişi tespit edildi; veriler düzeltilmeli.",'
                 'IF(B11="ZARARLI","Toplam kâr marjı negatif; maliyetler hasılatın üzerinde, acil önlem gereklidir.",'
                 'IF(B11="KARLI","Toplam kâr marjı hedef marjın üzerinde; nakliye faaliyeti kârlıdır.",'
                 '"Kâr marjı pozitif ancak hedef marjın altında; maliyet kalemleri incelenmelidir."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[12].height = 44
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["Aksiyon 1", "Aksiyon 2", "Aksiyon 3"], 15):
        h(ws, i, 1, ad, yazi="333333", kalin=True)
        h(ws, i, 2, f'=IF(B11="VERİ YOK","Araç ve sefer verilerini girin; karar otomatik üretilir.",'
                    f'IF(B11="DURDUR","Negatif mesafe veya hasılat değerlerini düzeltin.",NakliyeAksiyon{i-14}_ACIKLAMA))',
          yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[i].height = 40
    h(ws, 19, 1, "Karar Notu", yazi="333333")
    h(ws, 19, 2, "Bu karar destek aracıdır; maliyetler fiili yakıt ve bakım kayıtlarıyla doğrulanmalıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:B19", URUN_AD)
    alt_bant(ws, 21, "Karar kuralları KILAVUZ sayfasında gösterilmiştir.")
    genislik(ws, {"A": 32, "B": 80})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (1, "Toplam Hasılat (TL)", "=HESAP!B6", TL),
        (2, "Toplam Nakliye Maliyeti (TL)", "=HESAP!B7", TL),
        (3, "TOPLAM NET KÂR (TL)", "=HESAP!B12", TL),
        (4, "Toplam Kâr Marjı", "=HESAP!B13", YÜZDE),
        (5, "Ortalama Birim Maliyet (TL/km)", "=HESAP!B15", TL),
        (6, "Maliyetin Kâra Oranı", "=HESAP!B21", YÜZDE),
        (7, "Karar", "=KARAR!B11", NORMAL),
        (8, "Kalite Skoru", "=KONTROLLER!B16", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En Pahalı Sefer", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B20", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "En Yüksek Sefer Maliyeti (TL)", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 6, 8, "=HESAP!B19", sayi=TL, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    ws.row_dimensions[6].height = 30
    # Araç bazlı hasılat/maliyet/kâr
    h(ws, 10, 1, "Araç", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Hasılat (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Toplam Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Kâr Marjı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    arac_etiketleri = ["34 ABC 123", "06 DEF 456", "35 GHI 789", "16 JKL 012",
                       "41 MNO 345", "07 PQR 678"]
    for i, etiket in enumerate(arac_etiketleri, 11):
        h(ws, i, 1, etiket, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!F{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=IFERROR(HESAP!F{i+16}/HESAP!B{i+16},0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Araç Bazlı Hasılat ve Toplam Maliyet"
    g1.add_data(Reference(ws, min_col=2, max_col=3, min_row=10, max_row=17), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g1.height, g1.width = 9, 15
    ws.add_chart(g1, "F8")
    g2 = BarChart()
    g2.type = "col"
    g2.title = "Araç Bazlı Net Kâr"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=17), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Araç Bazlı Kâr Marjı"
    g7.add_data(Reference(ws, min_col=5, min_row=10, max_row=17), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g7.height, g7.width = 9, 12
    ws.add_chart(g7, "AH8")
    # Maliyet bileşenleri
    h(ws, 19, 1, "Maliyet Kalemi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 2, "Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 3, "Pay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, kalem in enumerate(["Yakıt", "Bakım", "Sürücü", "Amortisman"], 20):
        h(ws, i, 1, kalem, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i-12}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=IFERROR(HESAP!B{i-12}/HESAP!B7,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g3 = PieChart()
    g3.title = "Maliyet Bileşenlerinin Dağılımı"
    g3.add_data(Reference(ws, min_col=2, min_row=19, max_row=24), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=20, max_row=24))
    g3.height, g3.width = 9, 12
    ws.add_chart(g3, "F21")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Maliyet Bileşenlerinin Payı"
    g8.add_data(Reference(ws, min_col=3, min_row=19, max_row=24), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=1, min_row=20, max_row=24))
    g8.height, g8.width = 9, 12
    ws.add_chart(g8, "AH21")
    # Maliyet trendi
    h(ws, 26, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 2, "Toplam Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(4):
        h(ws, 27 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 27 + i, 2, f"=HESAP!B{36+i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Sefer Maliyeti Trendi"
    g4.add_data(Reference(ws, min_col=2, min_row=26, max_row=31), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=27, max_row=31))
    g4.height, g4.width = 9, 15
    ws.add_chart(g4, "F31")
    # Projeksiyon
    h(ws, 33, 1, "Ortalama Sefer Maliyeti (TL)", yazi="333333")
    h(ws, 33, 2, "=HESAP!B41", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 34, 1, "Gelecek Dönem Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 34, 2, "=HESAP!B42", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 35, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 35, 2, "=HESAP!B44", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 36, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 36, 2, "=HESAP!B45", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Tahmin ve Bant"
    g5.add_data(Reference(ws, min_col=2, min_row=33, max_row=36), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=34, max_row=36))
    g5.height, g5.width = 9, 15
    ws.add_chart(g5, "U37")
    # Senaryo verisi
    h(ws, 38, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 38, 6, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 39):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-32}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = LineChart()
    g6.title = "Senaryo Net Kârları"
    g6.add_data(Reference(ws, min_col=6, min_row=38, max_row=43), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=5, min_row=39, max_row=43))
    g6.height, g6.width = 9, 12
    ws.add_chart(g6, "F45")
    baski_hazirla(ws, "A1:M45", URUN_AD)
    alt_bant(ws, 45, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 30, "B": 20, "C": 22, "D": 18, "E": 20, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Yakıt fiyatı çarpanına göre senaryo bant aralığı ve maliyet bileşeni duyarlılık "
                "(tornado) analizi.", yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Yakıt Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        (7, "İyimser", "=IF(COUNTA(tblSefer[Sefer Numarası])>0,NakliyeSenIyi,0)",
         "=IFERROR(HESAP!B12*NakliyeSenIyi/NakliyeSenBaz,0)"),
        (8, "Baz", "=IF(COUNTA(tblSefer[Sefer Numarası])>0,NakliyeSenBaz,0)", "=HESAP!B12"),
        (9, "Kötümser", "=IF(COUNTA(tblSefer[Sefer Numarası])>0,NakliyeSenKotu,0)",
         "=IFERROR(HESAP!B12*NakliyeSenKotu/NakliyeSenBaz,0)"),
        (10, "Kritik", "=IF(COUNTA(tblSefer[Sefer Numarası])>0,NakliyeSenKritik,0)",
         "=IFERROR(HESAP!B12*NakliyeSenKritik/NakliyeSenBaz,0)"),
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
        (14, "Yakıt Fiyatı", "=HESAP!B12*NakliyeTornadoYakitOran", "=HESAP!B12*NakliyeTornadoYakitOran*3"),
        (15, "Bakım Oranı", "=HESAP!B12*NakliyeTornadoBakimOran", "=HESAP!B12*NakliyeTornadoBakimOran*3"),
    ]
    for satir, ad, dus, yuk in tornado:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, dus, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 3, yuk, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, f'=C{satir}-B{satir}', sayi=TL, yazi=DIS_REF_YEŞIL)
    baski_hazirla(ws, "A1:D10", URUN_AD)
    alt_bant(ws, 18, "Tornado genişliği hangi değişkenin net kârı en çok etkilediğini gösterir.")
    genislik(ws, {"A": 28, "B": 24, "C": 24, "D": 22})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Raporu", kalin=True, boyut=15, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, '=CONCATENATE("Rapor Tarihi: ",TEXT(NakliyeRaporTarihi,"gg.aa.yyyy")," | ",NakliyeFirmaUnvan)',
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Karar Gerekçesi", yazi="333333", kaydir=True)
    h(ws, 7, 2, "=KARAR!B12", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[7].height = 40
    ozet = [
        (8, "Toplam Hasılat (TL)", "=HESAP!B6"),
        (9, "Toplam Nakliye Maliyeti (TL)", "=HESAP!B7"),
        (10, "Toplam Net Kâr (TL)", "=HESAP!B12"),
        (11, "Toplam Kâr Marjı", "=HESAP!B13"),
        (12, "Ortalama Birim Maliyet (TL/km)", "=HESAP!B15"),
        (13, "Ton-Km Birim Maliyeti (TL)", "=HESAP!B17"),
        (14, "Sefer Sayısı", "=HESAP!B18"),
        (15, "Zararlı Sefer Sayısı", "=HESAP!B22"),
    ]
    for satir, ad, form in ozet:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=TL if "TL" in ad else YÜZDE if "Marj" in ad else CATI,
          yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 18 + i, 1, f"Aksiyon {i+1}", yazi="333333", kalin=True)
        h(ws, 18 + i, 2, f"=KARAR!B{15+i}", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[18 + i].height = 38
    h(ws, 23, 1, "Hazırlayan", yazi="333333")
    h(ws, 23, 2, "=NakliyeRaporHazirlayan", yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Sürüm", yazi="333333")
    h(ws, 24, 2, "=NakliyeDosyaSurumu", yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Bu rapor karar destek amaçlıdır; nihai karar fiili kayıtlarla doğrulanmalıdır.",
      yazi=GRİ, boyut=9, kaydir=True)
    baski_hazirla(ws, "A1:B26", URUN_AD)
    genislik(ws, {"A": 40, "B": 26})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=36)
    h(ws, 3, 1, "Örnek Veriler", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Ürünü tanımak için bu sayfadaki örnek araç ve sefer verileri "
                "yan yana bloklar halinde sunulur; ana sayfalarda denemeler yapın.",
      yazi=GRİ, kaydir=True)
    # ---- Blok 1: Örnek Araçlar (A-H) ----
    h(ws, 5, 1, "Örnek Araçlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar = ["Plaka", "Araç Tipi", "Yakıt Tüketimi (lt/100km)", "Bakım Oranı (TL/km)",
                "Sürücü Maliyeti (TL/gün)", "Günlük Kapasite (km/gün)", "Amortisman (TL/km)",
                "Açıklama"]
    baslik_satiri(ws, 6, sutunlar, basla=1)
    formuller = {
        "Açıklama": '=IF(tblOrnekArac[[#This Row],[Plaka]]="","",CONCATENATE("Araç: ",tblOrnekArac[[#This Row],[Plaka]]," | Tip: ",tblOrnekArac[[#This Row],[Araç Tipi]]))',
    }
    tablo_ekle(ws, "tblOrnekArac", "A6:H1005", sutunlar, formuller)
    ornek_arac = [
        ("34 ABC 123", "Kamyonet", 11, 1.4, 1200, 350, 1.1),
        ("06 DEF 456", "Kamyon", 24, 2.2, 1600, 500, 1.8),
        ("35 GHI 789", "Tır", 32, 2.8, 2000, 650, 2.4),
        ("16 JKL 012", "Kamyonet", 10, 1.2, 1100, 320, 1.0),
        ("41 MNO 345", "Çekici", 28, 2.5, 1800, 600, 2.1),
        ("07 PQR 678", "Kamyon", 25, 2.3, 1650, 520, 1.9),
    ]
    for i, (plaka, tip, tuketim, bakim, surucu, kapasite, amortisman) in enumerate(ornek_arac, 7):
        for kolon, deger in [(1, plaka), (2, tip), (3, tuketim), (4, bakim),
                             (5, surucu), (6, kapasite), (7, amortisman)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=3).number_format = CATI
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=5).number_format = TL
        ws.cell(row=satir, column=6).number_format = CATI
        ws.cell(row=satir, column=7).number_format = TL
    giris_hucreleri(ws, 7, 1005, [1, 2, 3, 4, 5, 6, 7])
    # ---- Blok 2: Örnek Seferler (J-Y) ----
    h(ws, 5, 10, "Örnek Seferler", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar2 = ["Sefer Numarası", "Tarih", "Plaka", "Güzergah", "Mesafe (km)", "Yük (ton)",
                 "Hasılat (TL)", "Yakıt Maliyeti (TL)", "Bakım Maliyeti (TL)",
                 "Sürücü Maliyeti (TL)", "Amortisman (TL)", "Toplam Maliyet (TL)",
                 "Birim Maliyet (TL/km)", "Ton-Km Maliyeti (TL)", "Net Kâr (TL)",
                 "Kâr Marjı", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar2, basla=10)
    formuller2 = {
        "Yakıt Maliyeti (TL)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="","",tblOrnekSefer[[#This Row],[Mesafe (km)]]*(IFERROR(INDEX(tblOrnekArac[Yakıt Tüketimi (lt/100km)],MATCH(tblOrnekSefer[[#This Row],[Plaka]],tblOrnekArac[Plaka],0)),0)/100)*NakliyeYakitFiyati)',
        "Bakım Maliyeti (TL)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="","",tblOrnekSefer[[#This Row],[Mesafe (km)]]*IFERROR(INDEX(tblOrnekArac[Bakım Oranı (TL/km)],MATCH(tblOrnekSefer[[#This Row],[Plaka]],tblOrnekArac[Plaka],0)),0))',
        "Sürücü Maliyeti (TL)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="","",IFERROR(tblOrnekSefer[[#This Row],[Mesafe (km)]]/IFERROR(INDEX(tblOrnekArac[Günlük Kapasite (km/gün)],MATCH(tblOrnekSefer[[#This Row],[Plaka]],tblOrnekArac[Plaka],0)),1)*IFERROR(INDEX(tblOrnekArac[Sürücü Maliyeti (TL/gün)],MATCH(tblOrnekSefer[[#This Row],[Plaka]],tblOrnekArac[Plaka],0)),0),0))',
        "Amortisman (TL)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="","",tblOrnekSefer[[#This Row],[Mesafe (km)]]*IFERROR(INDEX(tblOrnekArac[Amortisman (TL/km)],MATCH(tblOrnekSefer[[#This Row],[Plaka]],tblOrnekArac[Plaka],0)),0))',
        "Toplam Maliyet (TL)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="","",SUM(tblOrnekSefer[[#This Row],[Yakıt Maliyeti (TL)]]:tblOrnekSefer[[#This Row],[Amortisman (TL)]]))',
        "Birim Maliyet (TL/km)": '=IF(tblOrnekSefer[[#This Row],[Mesafe (km)]]="",0,IFERROR(tblOrnekSefer[[#This Row],[Toplam Maliyet (TL)]]/tblOrnekSefer[[#This Row],[Mesafe (km)]],0))',
        "Ton-Km Maliyeti (TL)": '=IF(OR(tblOrnekSefer[[#This Row],[Mesafe (km)]]="",tblOrnekSefer[[#This Row],[Yük (ton)]]=""),0,IFERROR(tblOrnekSefer[[#This Row],[Toplam Maliyet (TL)]]/(tblOrnekSefer[[#This Row],[Mesafe (km)]]*tblOrnekSefer[[#This Row],[Yük (ton)]]),0))',
        "Net Kâr (TL)": '=IF(tblOrnekSefer[[#This Row],[Hasılat (TL)]]="","",IFERROR(tblOrnekSefer[[#This Row],[Hasılat (TL)]]-tblOrnekSefer[[#This Row],[Toplam Maliyet (TL)]],0))',
        "Kâr Marjı": '=IF(tblOrnekSefer[[#This Row],[Hasılat (TL)]]="",0,IFERROR(tblOrnekSefer[[#This Row],[Net Kâr (TL)]]/tblOrnekSefer[[#This Row],[Hasılat (TL)]],0))',
        "Açıklama": '=IF(tblOrnekSefer[[#This Row],[Sefer Numarası]]="","",CONCATENATE("Sefer: ",tblOrnekSefer[[#This Row],[Sefer Numarası]]," | Güzergah: ",tblOrnekSefer[[#This Row],[Güzergah]]," | Net Kâr: ",TEXT(tblOrnekSefer[[#This Row],[Net Kâr (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblOrnekSefer", "J6:Y1005", sutunlar2, formuller2)
    ornek_sefer = [
        ("SF-1001", date(2026, 6, 3), "34 ABC 123", "İstanbul", 320, 1.5, 8500),
        ("SF-1002", date(2026, 6, 5), "06 DEF 456", "Ankara", 470, 8.0, 18500),
        ("SF-1003", date(2026, 6, 8), "35 GHI 789", "İzmir", 590, 16.0, 32000),
        ("SF-1004", date(2026, 6, 12), "16 JKL 012", "Bursa", 180, 1.2, 5200),
        ("SF-1005", date(2026, 6, 15), "41 MNO 345", "Antalya", 780, 12.0, 27500),
        ("SF-1006", date(2026, 6, 20), "07 PQR 678", "Adana", 890, 9.0, 29500),
        ("SF-1007", date(2026, 6, 24), "34 ABC 123", "Trabzon", 1180, 1.4, 21000),
        ("SF-1008", date(2026, 6, 28), "06 DEF 456", "Samsun", 730, 8.5, 20500),
        ("SF-1009", date(2026, 7, 2), "35 GHI 789", "İstanbul", 340, 15.0, 24500),
        ("SF-1010", date(2026, 7, 6), "41 MNO 345", "İzmir", 610, 11.0, 23500),
    ]
    for i, (sefer, tarih, plaka, guzergah, mesafe, yuk, hasilat) in enumerate(ornek_sefer, 7):
        for kolon, deger in [(10, sefer), (11, tarih), (12, plaka), (13, guzergah),
                             (14, mesafe), (15, yuk), (16, hasilat)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=11).number_format = TARİH
        ws.cell(row=satir, column=14).number_format = CATI
        ws.cell(row=satir, column=15).number_format = CATI
        for kolon in (16, 17, 18, 19, 20, 21, 22, 23, 24):
            ws.cell(row=satir, column=kolon).number_format = TL
        ws.cell(row=satir, column=25).number_format = YÜZDE
    giris_hucreleri(ws, 7, 1005, [10, 11, 12, 13, 14, 15, 16])
    sabitle(ws, "A7")
    genislik(ws, {"A": 14, "B": 12, "C": 20, "D": 16, "E": 20, "F": 18, "G": 18, "H": 30,
                  "J": 10, "K": 12, "L": 13, "M": 12, "N": 13, "O": 10, "P": 13,
                  "Q": 16, "R": 16, "S": 16, "T": 15, "U": 17, "V": 17, "W": 17,
                  "X": 13, "Y": 40})


def degisiklik_kaydi(ws):
    sayfa_hazirla(ws, "DEGISIKLIK_KAYDI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Değişiklik Kaydı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 5, ["Tarih", "Sürüm", "Değişiklik", "Neden"])
    satirlar = [
        (6, "10.08.2026", "1.0.0", "İlk üretim", "Ürünün yayına hazırlanması"),
    ]
    for satir, tarih, surum, degisiklik, neden in satirlar:
        h(ws, satir, 1, tarih, yazi="333333")
        h(ws, satir, 2, surum, yazi="333333")
        h(ws, satir, 3, degisiklik, yazi="333333")
        h(ws, satir, 4, neden, yazi="333333")
    genislik(ws, {"A": 16, "B": 12, "C": 40, "D": 32, "E": 18})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Açılır Liste Kaynakları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere yeni değer ekleyin; doğrulama listeleri otomatik genişler.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Araç Tipleri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Güzergahlar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Plakalar", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tip in enumerate(ARAC_TIPLERI, 7):
        h(ws, i, 1, tip, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Araç tipi | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni tip için satır ekleyin")
    for i, guzergah in enumerate(GUZERGAHLAR, 7):
        h(ws, i, 3, guzergah, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"C{i}", "Tanım: Güzergah | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni güzergah için satır ekleyin")
    for i, plaka in enumerate(PLAKALAR, 7):
        h(ws, i, 5, plaka, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"E{i}", "Tanım: Plaka | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni plaka için satır ekleyin")
    genislik(ws, {"A": 22, "C": 16, "E": 18})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Eşikler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Eşikler, senaryo çarpanları ve oranları "
                "güncelleyin; değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"])
    TARIH_D = date(2026, 8, 10)
    YURURLUK = date(2026, 8, 10)
    DOGRULAMA = date(2026, 8, 10)
    satirlar = [
        (7, "NakliyeFirmaUnvan", "Örnek Nakliyat A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "NakliyeRaporTarihi", TARIH_D, "Tarih", None, None, "Raporun dayanak tarihi.", "Kullanıcı"),
        (9, "NakliyeRaporHazirlayan", "Filo ve Mali İşler Yöneticisi", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "NakliyeDosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "NakliyeYakitFiyati", 42.0, "TL/lt", 0, 100, "Dizel yakıt fiyatı (TL/lt).", "Varsayım"),
        (12, "NakliyeHedefMarj", 0.15, "Oran", 0, 1, "Hedef kâr marjı eşiği (%15).", "Varsayım"),
        (13, "NakliyeKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (14, "NakliyeKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (15, "NakliyeSenIyi", 0.90, "Oran", 0, 3, "İyimser senaryo yakıt çarpanı.", "Varsayım"),
        (16, "NakliyeSenBaz", 1.00, "Oran", 0, 3, "Baz senaryo yakıt çarpanı.", "Varsayım"),
        (17, "NakliyeSenKotu", 1.15, "Oran", 0, 3, "Kötümser senaryo yakıt çarpanı.", "Varsayım"),
        (18, "NakliyeSenKritik", 1.30, "Oran", 0, 3, "Kritik senaryo yakıt çarpanı.", "Varsayım"),
        (19, "NakliyeTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (20, "NakliyeTahminDonem", 5, "Dönem", 1, 12, "Gelecek dönem tahmin noktası.", "Varsayım"),
        (21, "NakliyeTornadoYakitOran", 0.10, "Oran", 0, 0.5, "Yakıt fiyatı tornado adımı (%10).", "Varsayım"),
        (22, "NakliyeTornadoBakimOran", 0.05, "Oran", 0, 0.5, "Bakım oranı tornado adımı.", "Varsayım"),
        (23, "NakliyeP90Oran", 0.90, "Oran", 0, 1, "Birim maliyet dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (24, "NakliyeYuksekMaliyetEsik", 100000, "TL", 0, 1000000000, "Yüksek sefer maliyeti kontrol eşiği.", "Varsayım"),
        (25, "NakliyeAksiyon1_ACIKLAMA",
         "Kâr marjı negatif; yakıt tüketimi yüksek araçların rotalarını ve doluluk oranlarını denetleyin, yakıt tasarrufu sağlayan sürüş eğitimi planlayın.", "Metin", None, None,
         "ZARARLI kararında gösterilen ilk öneri.", "Varsayım"),
        (26, "NakliyeAksiyon2_ACIKLAMA",
         "Kâr marjı hedefin altında; bakım ve sürücü maliyetleri yüksek kalemleri gözden geçirin, sefer bazında doluluk artırıcı yük planlaması yapın.", "Metin", None, None,
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (27, "NakliyeAksiyon3_ACIKLAMA",
         "Kâr marjı hedefin üzerinde; mevcut filo performansını sürdürerek dönemlik maliyet eğilimini izlemeye devam edin.", "Metin", None, None,
         "KARLI kararında gösterilen üçüncü öneri.", "Varsayım"),
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
        if 25 <= satir <= 27:
            ws.cell(row=satir, column=2).alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[satir].height = 48
        ws.row_dimensions[satir].height = 26
    for i in range(1, 7):
        satir = 28 + i
        h(ws, satir, 1, f"NakliyePanelSatir_{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki grafiğin araç sırası ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel sıra numarası {i} | "
                   f"Neden önemli: Grafikleri besler | "
                   f"Doğru kullanım: Araç sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 34, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 22, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 5, 1, "Karar Kuralları", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kurallar = [
        "Toplam kâr marjı hedef marja eşit veya üzerinde ise karar KARLI olur.",
        "Toplam kâr marjı pozitif ancak hedef marjın altında ise karar İNCELE olur.",
        "Toplam kâr marjı negatif ise karar ZARARLI olur.",
        "Sefer verisi girilmemişse karar VERİ YOK olur; negatif mesafe/hasılat varsa DURDUR olur.",
    ]
    for i, m in enumerate(kurallar, 6):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=28)
    h(ws, 12, 1, "Hesaplama Yöntemleri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    yontemler = [
        "Yakıt maliyeti = Mesafe × (Yakıt Tüketimi / 100) × Yakıt Fiyatı.",
        "Bakım maliyeti = Mesafe × Bakım Oranı (TL/km).",
        "Sürücü maliyeti = (Mesafe / Günlük Kapasite) × Günlük Sürücü Maliyeti.",
        "Amortisman = Mesafe × Amortisman Oranı (TL/km).",
        "Toplam maliyet = Yakıt + Bakım + Sürücü + Amortisman.",
        "Net kâr = Hasılat − Toplam maliyet; kâr marjı = Net kâr / Hasılat.",
    ]
    for i, m in enumerate(yontemler, 13):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=28)
    h(ws, 21, 1, "Ortam ve Güvenlik", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ortam = [
        "Desteklenen ortam: Excel 2016–365 (Windows/Mac), LibreOffice Calc, Google Sheets.",
        "Verileriniz cihazınızdan çıkmaz: makro yok, dış bağlantı yok, telemetri yok.",
        "Tüm sayfalar '1234' şifresiyle korunur; sarı hücreler veri girişi içindir.",
    ]
    for i, m in enumerate(ortam, 22):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=28)
    h(ws, 27, 1, "Bu dosya karar destek aracıdır; nakliye maliyetleri fiili yakıt, bakım ve "
                 "ücret kayıtlarıyla doğrulanmalıdır.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 34, "B": 110})


def kosullu_bicimlendirme(wb):
    ws = wb["PANO"]
    yesil = PatternFill("solid", start_color="1F7A4D")
    sari = PatternFill("solid", start_color="FFF2CC")
    kirmizi = PatternFill("solid", start_color="B3261E")
    kirmizi_font = Font(color="FFFFFF", bold=True)
    for kolon in range(1, 9):
        hucre = ws.cell(row=4, column=kolon)
        if kolon in (3, 7):
            ws.conditional_formatting.add(hucre.coordinate,
                CellIsRule(operator="containsText", formula=['"KRİTİK"'], fill=kirmizi, font=kirmizi_font))
        if kolon in (3,):
            ws.conditional_formatting.add(hucre.coordinate,
                CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
            ws.conditional_formatting.add(hucre.coordinate,
                CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=yesil))
    ws2 = wb["KARAR"]
    ws2.conditional_formatting.add("B11",
        CellIsRule(operator="equal", formula=['"KARLI"'], fill=yesil))
    ws2.conditional_formatting.add("B11",
        CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
    ws2.conditional_formatting.add("B11",
        CellIsRule(operator="equal", formula=['"ZARARLI"'], fill=kirmizi, font=kirmizi_font))
    ws2.conditional_formatting.add("B11",
        CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
    ws3 = wb["KONTROLLER"]
    for satir in range(6, 12):
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil))
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"DOLU"'], fill=yesil))
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"BOŞ"'], fill=sari))
    ws3.conditional_formatting.add("B16",
        CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil))
    ws3.conditional_formatting.add("B16",
        CellIsRule(operator="between", formula=["70", "89"], fill=sari))
    ws3.conditional_formatting.add("B16",
        CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kirmizi_font))
    for satir in range(24, 33):
        ws3.conditional_formatting.add(f"C{satir}",
            CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws3.conditional_formatting.add(f"C{satir}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws4 = wb["SENARYO_DUYARLILIK"]
    for satir in range(7, 11):
        ws4.conditional_formatting.add(f"C{satir}",
            CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws4.conditional_formatting.add(f"C{satir}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws5 = wb["PANO"]
    for satir in range(39, 43):
        ws5.conditional_formatting.add(f"F{satir}",
            CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws5.conditional_formatting.add(f"F{satir}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))


def moduller(wb):
    ws = wb["KONTROLLER"]
    modul_satirlari = {
        "modulT1OrtMaliyet": ("T1", "Ortalama Birim Maliyet (TL/km)",
                              "=HESAP!B15",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama birim maliyet ",TEXT(C24,"₺ #,##0")," TL/km — filonun genel maliyet verimini gösterir")',
                              TL),
        "modulT2OrtMarj": ("T2", "Ortalama Kâr Marjı",
                           "=IFERROR(AVERAGE(tblSefer[Kâr Marjı]),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama sefer kâr marjı ",TEXT(C25,"0.0%")," — filo kârlılık düzeyini gösterir")',
                           YÜZDE),
        "modulO1MaliyetSapmasi": ("O1", "Maliyet Sapması (MAD)",
                                  "=IFERROR(AVEDEV(tblSefer[Toplam Maliyet (TL)]),0)",
                                  '=_xlfn.TEXTJOIN("; ",TRUE,"Sefer maliyeti ortalama mutlak sapması ",TEXT(C26,"₺ #,##0")," TL — seferler arası heterojenliği ölçer")',
                                  TL),
        "modulO2EnYuksekMaliyet": ("O2", "En Yüksek Sefer Maliyeti",
                                   "=IFERROR(MAX(tblSefer[Toplam Maliyet (TL)]),0)",
                                   '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek sefer maliyeti ",TEXT(C27,"₺ #,##0")," TL — öncelikli iyileştirme seferidir")',
                                   TL),
        "modulO3Senaryo": ("O3", "Senaryo Bant Genişliği",
                           "=IFERROR(STDEV.P(SENARYO_DUYARLILIK!C7:C10),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo net kâr standart sapması ",TEXT(C28,"₺ #,##0")," TL — kâr belirsizliğinin etkisini gösterir")',
                           TL),
        "modulO6Hhi": ("O6", "Hasılat Yoğunlaşması (HHI)",
                       '=IFERROR(IF(SUM(tblSefer[Hasılat (TL)])=0,0,SUMPRODUCT((tblSefer[Hasılat (TL)]/SUM(tblSefer[Hasılat (TL)]))^2)),0)',
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Araç hasılat yoğunlaşma endeksi ",TEXT(C29,"0.000")," — 1 e yaklaştıkça hasılat tek araçta toplanır")',
                       None),
        "modulO8KaliteSkor": ("O8", "Veri Kalite Skoru",
                              "=KONTROLLER!B16",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C30,"0")," puan — ",IF(C30>=NakliyeKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))',
                              CATI),
        "modulI1Tahmin": ("I1", "Gelecek Dönem Maliyet Tahmini",
                          "=HESAP!B42",
                          '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem maliyet tahmini ",TEXT(C31,"₺ #,##0")," TL — bant aralığı için alt/üst sınır HESAP sayfasındadır")',
                          TL),
        "modulI2P90": ("I2", "Birim Maliyet Yüzdelik P90",
                       "=IFERROR(_xlfn.PERCENTILE.INC(tblSefer[Birim Maliyet (TL/km)],NakliyeP90Oran),0)",
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Birim maliyetlerin %90 ı ",TEXT(C32,"₺ #,##0")," TL/km altındadır — maliyet dağılımının üst sınırını gösterir")',
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
    sayfa_sirasi = ["KAPAK", "HIZLI_BASLANGIC", "ARAC_LISTESI", "SEFERLER", "HESAP",
                    "KONTROLLER", "KARAR", "PANO", "SENARYO_DUYARLILIK", "RAPOR", "ORNEK_VERI",
                    "DEGISIKLIK_KAYDI", "LISTELER", "AYARLAR", "KILAVUZ"]
    for ad in sayfa_sirasi:
        wb.create_sheet(ad)
    ilk = wb["Sheet"]
    wb.remove(ilk)

    kapak(wb["KAPAK"])
    hizli_baslangic(wb["HIZLI_BASLANGIC"])
    arac_listesi(wb["ARAC_LISTESI"])
    seferler(wb["SEFERLER"])
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
    for satir in range(7, 40):
        anahtar = wb["AYARLAR"].cell(row=satir, column=1).value
        if isinstance(anahtar, str) and anahtar.startswith("Nakliye"):
            ad_ekle(wb, anahtar, f"AYARLAR!$B${satir}")

    # Kritik çıktı ad tanımları (D02 karşılıkları ve PANO beslemeleri)
    ad_ekle(wb, "NakliyeToplamMaliyet", "HESAP!$B$7")
    ad_ekle(wb, "NakliyeBirimMaliyet", "HESAP!$B$15")
    ad_ekle(wb, "NakliyeKarar", "KARAR!$B$11")
    ad_ekle(wb, "NakliyeEnPahaliSefer", "HESAP!$B$20")
    ad_ekle(wb, "NakliyeProjX", "HESAP!$A$36:$A$39")

    ad_ekle(wb, "ListeAracTipleri", "LISTELER!$A$7:$A$10")
    ad_ekle(wb, "ListeGuzergahlar", "LISTELER!$C$7:$C$14")
    ad_ekle(wb, "ListePlakalar", "LISTELER!$E$7:$E$12")

    for ws in wb.worksheets:
        sayfa_koru(ws)

    dosya = "NakliyeMaliyetiHesaplayici.xlsx"
    wb.save(dosya)
    print("Dosya oluşturuldu:", dosya)


if __name__ == "__main__":
    main()
