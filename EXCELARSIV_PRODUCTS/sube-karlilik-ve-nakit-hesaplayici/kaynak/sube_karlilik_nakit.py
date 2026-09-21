"""Şube Kârlılık ve Nakit Hesaplayıcı — Excel Üretim Mandası v4.0 uyumlu üretim betiği."""
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

URUN_AD = "Şube Kârlılık ve Nakit Hesaplayıcı"
SURUM = "1.0.0"
RENK = "B08948"

SEHIRLER = ["İstanbul", "Ankara", "İzmir", "Bursa", "Antalya", "Adana", "Trabzon", "Samsun"]
KALEMLER = ["Satış", "Hizmet", "Kira Gideri", "Personel", "Elektrik", "Su", "Malzeme", "Diğer"]
SUBE_ADLARI = ["İstanbul Merkez", "Ankara Şubesi", "İzmir Şubesi", "Bursa Şubesi",
               "Antalya Şubesi", "Adana Şubesi"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Şube bazında gelir, gider ve nakit akışını toplayarak şube kârlılığını ve "
                "nakit pozisyonunu hesaplayın; kâr marjını eşiklerle karşılaştırıp "
                "iyileştirme önceliklerini gerekçesiyle raporlayın.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Şube bazlı gelir ve gider kayıtlarının canlı formülle kârlılığa dönüştürülmesi",
        "Tahsilat ve ödeme kayıtlarından nakit akışı ve nakit pozisyonunun hesaplanması",
        "Kâr marjını eşiklerle karşılaştırarak KARLI / İNCELE / ZARARLI kararının üretilmesi",
        "Gider kalemleri, gelir yoğunlaşması, anomali, senaryo ve duyarlılık analizi",
        "Gelecek dönem nakit tahmini ve veri kalite skoruyla karar kapısına güven katmanı",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Şube bazında kârlılığı ve nakit pozisyonunu izlemek isteyen işletme sahipleri",
        "Şube performansını karşılaştıran bölge ve operasyon yöneticileri",
        "Şube gelir/gider ve nakit akışını raporlayan mali işler yöneticileri",
        "Müşterisine şube kârlılık ve nakit analizi öneren mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 20, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. SUBE_LISTESI, GELIR_GIDER ve NAKIT_AKISI sayfalarına verilerinizi girin. "
                 "3. PANO ve KARAR sayfalarından şube kârlılık kararını izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A21:P21")
    h(ws, 23, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A23:P23")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Şube listesini girin",
         "SUBE_LISTESI sayfasındaki sarı hücrelere şube adını, şehri, açılış nakitini ve "
         "hedef kâr marjını girin."),
        ("Adım 2 — Gelir ve giderleri girin",
         "GELIR_GIDER sayfasındaki sarı hücrelere tarih, şube, kalem, tür (Gelir/Gider) ve "
         "tutarı girin; işaretli tutar otomatik hesaplanır."),
        ("Adım 3 — Nakit akışını girin",
         "NAKIT_AKISI sayfasındaki sarı hücrelere tahsilat ve ödeme kayıtlarını girin."),
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
        "Gelir kalemini tür olarak Gider seçmek (işaretli tutar ters işaretli olur)",
        "Şube adını SUBE_LISTESI ile birebir aynı yazmamak (şube analizi eşleşmez)",
        "Tahsilatı Ödeme olarak kaydetmek (nakit pozisyonu hatalı küçülür)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 16, 1, "Sık sorulan sorular", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sss = [
        "Soru: Kâr marjı nasıl hesaplanır? Cevap: Net kârın (gelir − gider) gelire bölünmesiyle otomatik üretilir.",
        "Soru: Nakit pozisyonu nasıl bulunur? Cevap: Açılış nakitine tahsilatlar eklenir, ödemeler çıkarılır.",
        "Soru: Yeni bir şube nasıl eklenir? Cevap: SUBE_LISTESI ve LISTELER sayfalarındaki sarı hücrelere ekleyin.",
        "Soru: Kalem listesinde aradığım yoksa? Cevap: LISTELER sayfasındaki sarı hücreye yeni kalemi ekleyin.",
    ]
    for i, m in enumerate(sss, 17):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
        ws.row_dimensions[i].height = 32
    h(ws, 22, 1, "Bu dosya karar destek aracıdır; şube kârlılığı fiili gelir, gider ve nakit "
                 "kayıtlarıyla doğrulanmalıdır.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def sube_listesi(ws):
    sayfa_hazirla(ws, "SUBE_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şube Listesi ve Nakit Parametreleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Şube bazlı gelir, gider ve nakit "
                "analizi bu listeden beslenir.", yazi=GRİ, kaydir=True)
    sutunlar = ["Şube Adı", "Şehir", "Açılış Nakit (TL)", "Hedef Kâr Marjı", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Açıklama": '=IF(tblSube[[#This Row],[Şube Adı]]="","",CONCATENATE("Şube: ",tblSube[[#This Row],[Şube Adı]]," | Şehir: ",tblSube[[#This Row],[Şehir]]," | Açılış Nakit: ",TEXT(tblSube[[#This Row],[Açılış Nakit (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblSube", "A5:E1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeSehirler", "B6:B1004",
              baslik="Şehir", mesaj="Listeden şehri seçin.",
              hata_baslik="Geçersiz şehir", hata_mesaj="Listede olmayan şehir giremezsiniz.")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Açılış Nakit", mesaj="0 veya daha büyük değer girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Hedef Kâr Marjı", mesaj="0 ile 1 arasında oran girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1 arasında olmalı.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = YÜZDE
    ornek = [
        ("İstanbul Merkez", "İstanbul", 250000, 0.15),
        ("Ankara Şubesi", "Ankara", 180000, 0.12),
        ("İzmir Şubesi", "İzmir", 220000, 0.13),
        ("Bursa Şubesi", "Bursa", 120000, 0.10),
        ("Antalya Şubesi", "Antalya", 160000, 0.14),
        ("Adana Şubesi", "Adana", 110000, 0.11),
    ]
    for i, (ad, sehir, nakit, marj) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = ad
        ws.cell(row=i, column=2).value = sehir
        ws.cell(row=i, column=3).value = nakit
        ws.cell(row=i, column=4).value = marj
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4])
    sabitle(ws, "A6")
    genislik(ws, {"A": 20, "B": 14, "C": 18, "D": 16, "E": 60})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def gelir_gider(ws):
    sayfa_hazirla(ws, "GELIR_GIDER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Gelir ve Gider Kayıtları", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Her kaydın tarih, şube, kalem, tür ve tutar bilgisini girin; işaretli "
                "tutar ve kârlılık katkısı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Şube", "Kalem", "Tür", "Tutar (TL)", "İşaretli Tutar (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "İşaretli Tutar (TL)": '=IF(tblGelirGider[[#This Row],[Tutar (TL)]]="","",IF(tblGelirGider[[#This Row],[Tür]]="Gider",-tblGelirGider[[#This Row],[Tutar (TL)]],tblGelirGider[[#This Row],[Tutar (TL)]]))',
        "Açıklama": '=IF(tblGelirGider[[#This Row],[Tarih]]="","",CONCATENATE("Kayıt: ",TEXT(tblGelirGider[[#This Row],[Tarih]],"gg.aa.yyyy")," | Şube: ",tblGelirGider[[#This Row],[Şube]]," | ",tblGelirGider[[#This Row],[Kalem]]," | ",tblGelirGider[[#This Row],[Tür]]," | Tutar: ",TEXT(tblGelirGider[[#This Row],[Tutar (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblGelirGider", "A5:G1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeSubeler", "B6:B1004",
              baslik="Şube", mesaj="Listeden şubeyi seçin.",
              hata_baslik="Geçersiz şube", hata_mesaj="Şube listesinde olmayan ad giremezsiniz.")
    dogrulama(ws, "list", "ListeKalemler", "C6:C1004",
              baslik="Kalem", mesaj="Listeden kalemi seçin.",
              hata_baslik="Geçersiz kalem", hata_mesaj="Listede olmayan kalem giremezsiniz.")
    dogrulama(ws, "list", "ListeTurler", "D6:D1004",
              baslik="Tür", mesaj="Gelir veya Gider seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Yalnızca Gelir veya Gider girebilirsiniz.")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Tutar", mesaj="0 veya daha büyük değer girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "date", None, "A6:A1004",
              baslik="Tarih", mesaj="Kayıt tarihini GG.AA.YYYY biçiminde girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Geçerli bir tarih girin.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=5).number_format = TL
        ws.cell(row=satir, column=6).number_format = TL
    ornek = [
        (date(2026, 6, 5), "İstanbul Merkez", "Satış", "Gelir", 480000),
        (date(2026, 6, 5), "İstanbul Merkez", "Kira Gideri", "Gider", 45000),
        (date(2026, 6, 8), "Ankara Şubesi", "Satış", "Gelir", 310000),
        (date(2026, 6, 8), "Ankara Şubesi", "Personel", "Gider", 120000),
        (date(2026, 6, 10), "İzmir Şubesi", "Hizmet", "Gelir", 265000),
        (date(2026, 6, 10), "İzmir Şubesi", "Elektrik", "Gider", 22000),
        (date(2026, 6, 12), "Bursa Şubesi", "Satış", "Gelir", 185000),
        (date(2026, 6, 12), "Bursa Şubesi", "Malzeme", "Gider", 75000),
        (date(2026, 6, 15), "Antalya Şubesi", "Satış", "Gelir", 240000),
        (date(2026, 6, 15), "Antalya Şubesi", "Personel", "Gider", 90000),
        (date(2026, 6, 18), "Adana Şubesi", "Hizmet", "Gelir", 155000),
        (date(2026, 6, 18), "Adana Şubesi", "Su", "Gider", 18000),
        (date(2026, 6, 20), "İstanbul Merkez", "Satış", "Gelir", 460000),
        (date(2026, 6, 20), "Ankara Şubesi", "Kira Gideri", "Gider", 40000),
    ]
    for i, (tarih, sube, kalem, tur, tutar) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = tarih
        ws.cell(row=i, column=2).value = sube
        ws.cell(row=i, column=3).value = kalem
        ws.cell(row=i, column=4).value = tur
        ws.cell(row=i, column=5).value = tutar
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5])
    sabitle(ws, "A6")
    genislik(ws, {"A": 12, "B": 20, "C": 14, "D": 10, "E": 13, "F": 17, "G": 60})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def nakit_akisi(ws):
    sayfa_hazirla(ws, "NAKIT_AKISI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nakit Akışı Kayıtları", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Tahsilat ve ödeme kayıtlarını girin; işaretli tutar ve nakit pozisyonu "
                "otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Şube", "İşlem Türü", "Tutar (TL)", "İşaretli Tutar (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "İşaretli Tutar (TL)": '=IF(tblNakit[[#This Row],[Tutar (TL)]]="","",IF(tblNakit[[#This Row],[İşlem Türü]]="Ödeme",-tblNakit[[#This Row],[Tutar (TL)]],tblNakit[[#This Row],[Tutar (TL)]]))',
        "Açıklama": '=IF(tblNakit[[#This Row],[Tarih]]="","",CONCATENATE("İşlem: ",TEXT(tblNakit[[#This Row],[Tarih]],"gg.aa.yyyy")," | Şube: ",tblNakit[[#This Row],[Şube]]," | ",tblNakit[[#This Row],[İşlem Türü]]," | Tutar: ",TEXT(tblNakit[[#This Row],[Tutar (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblNakit", "A5:F1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeSubeler", "B6:B1004",
              baslik="Şube", mesaj="Listeden şubeyi seçin.",
              hata_baslik="Geçersiz şube", hata_mesaj="Şube listesinde olmayan ad giremezsiniz.")
    dogrulama(ws, "list", "ListeIslemTurleri", "C6:C1004",
              baslik="İşlem Türü", mesaj="Tahsilat veya Ödeme seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Yalnızca Tahsilat veya Ödeme girebilirsiniz.")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Tutar", mesaj="0 veya daha büyük değer girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "date", None, "A6:A1004",
              baslik="Tarih", mesaj="İşlem tarihini GG.AA.YYYY biçiminde girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Geçerli bir tarih girin.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=5).number_format = TL
    ornek = [
        (date(2026, 6, 5), "İstanbul Merkez", "Tahsilat", 480000),
        (date(2026, 6, 6), "İstanbul Merkez", "Ödeme", 45000),
        (date(2026, 6, 8), "Ankara Şubesi", "Tahsilat", 310000),
        (date(2026, 6, 9), "Ankara Şubesi", "Ödeme", 120000),
        (date(2026, 6, 10), "İzmir Şubesi", "Tahsilat", 265000),
        (date(2026, 6, 11), "İzmir Şubesi", "Ödeme", 22000),
        (date(2026, 6, 12), "Bursa Şubesi", "Tahsilat", 185000),
        (date(2026, 6, 13), "Bursa Şubesi", "Ödeme", 75000),
        (date(2026, 6, 15), "Antalya Şubesi", "Tahsilat", 240000),
        (date(2026, 6, 16), "Antalya Şubesi", "Ödeme", 90000),
        (date(2026, 6, 18), "Adana Şubesi", "Tahsilat", 155000),
        (date(2026, 6, 19), "Adana Şubesi", "Ödeme", 18000),
        (date(2026, 6, 20), "İstanbul Merkez", "Tahsilat", 460000),
        (date(2026, 6, 22), "Ankara Şubesi", "Ödeme", 40000),
    ]
    for i, (tarih, sube, tur, tutar) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = tarih
        ws.cell(row=i, column=2).value = sube
        ws.cell(row=i, column=3).value = tur
        ws.cell(row=i, column=4).value = tutar
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4])
    sabitle(ws, "A6")
    genislik(ws, {"A": 12, "B": 20, "C": 12, "D": 13, "E": 17, "F": 60})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şube Kârlılık ve Nakit Hesaplama Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Şube, gelir/gider ve nakit verilerinden canlı üretilir; toplam kârlılık, "
                "marj ve nakit pozisyonu hesaplanır.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "TOPLAM GELİR (TL)", yazi="333333", kalin=True)
    h(ws, 6, 2, '=IFERROR(SUMIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Tür],"Gelir"),0)', sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "TOPLAM GİDER (TL)", yazi="333333", kalin=True)
    h(ws, 7, 2, '=IFERROR(SUMIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Tür],"Gider"),0)', sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 8, 1, "Toplam Net Kâr (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 8, 2, "=B6-B7", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 9, 1, "Toplam Kâr Marjı", yazi="333333")
    h(ws, 9, 2, "=IFERROR(B8/B6,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Toplam Tahsilat (TL)", yazi="333333")
    h(ws, 10, 2, '=IFERROR(SUMIFS(tblNakit[Tutar (TL)],tblNakit[İşlem Türü],"Tahsilat"),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Toplam Ödeme (TL)", yazi="333333")
    h(ws, 11, 2, '=IFERROR(SUMIFS(tblNakit[Tutar (TL)],tblNakit[İşlem Türü],"Ödeme"),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Nakit Akışı Net (TL)", yazi="333333")
    h(ws, 12, 2, "=B10-B11", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Toplam Nakit Pozisyonu (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 13, 2, "=IFERROR(SUM(tblSube[Açılış Nakit (TL)])+B12,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 14, 1, "Kayıt Sayısı", yazi="333333")
    h(ws, 14, 2, "=COUNT(tblGelirGider[Tarih])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Ortalama Kayıt Etkisi (TL)", yazi="333333")
    h(ws, 15, 2, "=IFERROR(AVERAGE(tblGelirGider[İşaretli Tutar (TL)]),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Şube Sayısı", yazi="333333")
    h(ws, 16, 2, "=COUNTA(tblSube[Şube Adı])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "En Yüksek Gider (TL)", yazi="333333")
    h(ws, 17, 2, '=IFERROR(_xlfn.MAXIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Tür],"Gider"),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Zararlı Kayıt Sayısı", yazi="333333")
    h(ws, 18, 2, '=COUNTIF(tblGelirGider[İşaretli Tutar (TL)],"<0")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Nakit İşlem Sayısı", yazi="333333")
    h(ws, 19, 2, "=COUNT(tblNakit[Tarih])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Gelir Kalemi Sayısı", yazi="333333")
    h(ws, 20, 2, '=COUNTIF(tblGelirGider[Tür],"Gelir")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Gider Kalemi Sayısı", yazi="333333")
    h(ws, 21, 2, '=COUNTIF(tblGelirGider[Tür],"Gider")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Nakit Oranı (Tahsilat/Ödeme)", yazi="333333")
    h(ws, 22, 2, "=IFERROR(B10/B11,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    # Şube bazlı panel verisi (grafikler için)
    h(ws, 26, 1, "Şube Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 2, "Gelir (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 4, "Gider (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 6, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 8, "Kâr Marjı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(6):
        h(ws, 27 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 27 + i, 2, f'=IFERROR(SUMIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Şube],INDEX(tblSube[Şube Adı],SubePanelSatir_{i+1}),tblGelirGider[Tür],"Gelir"),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 27 + i, 4, f'=IFERROR(SUMIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Şube],INDEX(tblSube[Şube Adı],SubePanelSatir_{i+1}),tblGelirGider[Tür],"Gider"),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 27 + i, 6, f'=IFERROR(B{27+i}-D{27+i},0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 27 + i, 8, f'=IFERROR(F{27+i}/B{27+i},0)', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    # Gider kalemleri
    h(ws, 34, 1, "Gider Kalemi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 34, 2, "Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, kalem in enumerate(["Kira Gideri", "Personel", "Elektrik", "Su", "Malzeme", "Diğer"], 35):
        h(ws, i, 1, kalem, yazi=GRİ)
        h(ws, i, 2, f'=IFERROR(SUMIFS(tblGelirGider[Tutar (TL)],tblGelirGider[Tür],"Gider",tblGelirGider[Kalem],A{i}),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    # Nakit trendi yardımcı serisi
    h(ws, 42, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 42, 2, "Nakit Akışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    nakit_serisi = [("1", "=IFERROR(INDEX(tblNakit[İşaretli Tutar (TL)],1),0)"),
                    ("2", "=IFERROR(INDEX(tblNakit[İşaretli Tutar (TL)],2),0)"),
                    ("3", "=IFERROR(INDEX(tblNakit[İşaretli Tutar (TL)],3),0)"),
                    ("4", "=IFERROR(INDEX(tblNakit[İşaretli Tutar (TL)],4),0)")]
    for i, (donem, form) in enumerate(nakit_serisi, 43):
        h(ws, i, 1, donem, sayi=CATI, yazi=GRİ)
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 48, 1, "Ortalama Nakit Akışı (TL)", yazi="333333")
    h(ws, 48, 2, "=IFERROR(AVERAGE(B43:B46),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 49, 1, "Gelecek Dönem Nakit Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 49, 2, "=IFERROR(FORECAST.LINEAR(SubeTahminDonem,B43:B46,SubeProjX)-B46,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 50, 1, "Nakit Std Sapması (TL)", yazi="333333")
    h(ws, 50, 2, "=IFERROR(STDEV.P(B43:B46),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 51, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 51, 2, "=IFERROR(B48-B50*SubeTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 52, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 52, 2, "=IFERROR(B48+B50*SubeTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 54, "Motor, şube ve kayıt verilerinden canlı beslenir; tüm eşikler AYARLAR'dan okunur.")
    genislik(ws, {"A": 34, "B": 40, "C": 24, "D": 24, "E": 24, "F": 22, "G": 22, "H": 20})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Veri Kalite ve Kontrol Merkezi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Girdilerin kalitesini ölçer; eksik, negatif ve uç değerleri raporlar.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    baslik_satiri(ws, 5, ["Kontrol", "Sonuç", "Durum"])
    kontroller = [
        ("Şube girildi mi?", '=IF(COUNTA(tblSube[Şube Adı])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Gelir/gider kaydı girildi mi?", '=IF(COUNTA(tblGelirGider[Tarih])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Nakit kaydı girildi mi?", '=IF(COUNTA(tblNakit[Tarih])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Negatif tutar girişi var mı?", '=IF(COUNTIF(tblGelirGider[Tutar (TL)],"<0")+COUNTIF(tblNakit[Tutar (TL)],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, "B3261E"),
        ("Zararlı kayıt var mı?", '=IF(COUNTIF(tblGelirGider[İşaretli Tutar (TL)],"<0")=0,"TEMİZ","GİDER VAR")', NORMAL, "B3261E"),
        ("Gelir kaydı yeterli mi?", '=IF(COUNTIF(tblGelirGider[Tür],"Gelir")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
        ("Gider kaydı yeterli mi?", '=IF(COUNTIF(tblGelirGider[Tür],"Gider")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
    ]
    for i, (ad, form, durum1, durum2) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "AUTOMATİK", yazi=GRİ)
    h(ws, 15, 1, "Veri Kalite Skoru", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Doluluk Oranı (%)", yazi="333333")
    h(ws, 16, 2, '=IFERROR(MIN(100,(COUNTA(tblSube[Şube Adı])+COUNTA(tblGelirGider[Tarih])+COUNTA(tblNakit[Tarih]))*100/SubeYeterliVeri),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Veri Kalite Skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, '=MIN(100,MAX(0,B16))', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Kalite Durumu", yazi="333333")
    h(ws, 18, 2, '=IF(B17>=SubeKaliteIyi,"YÜKSEK",IF(B17>=SubeKaliteOrta,"ORTA","DÜŞÜK"))', yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Canlı Formül Sayacı", yazi="333333")
    h(ws, 19, 2, '=IF(COUNTIF(tblGelirGider[Açıklama],"")=0,COUNTA(tblGelirGider[Açıklama]),0)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 21, "Tüm kontroller canlı formüllerle üretilir; sahte geçiş yoktur.")
    genislik(ws, {"A": 38, "B": 60, "C": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Şube Kârlılık Karar Kapısı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Kâr marjı, zararlı kayıt durumu ve nakit pozisyonuna göre karar üretir.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Toplam Kâr Marjı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B9", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Toplam Net Kâr (TL)", yazi="333333")
    h(ws, 7, 2, "=HESAP!B8", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Hedef Kâr Marjı", yazi="333333")
    h(ws, 8, 2, "=IF(COUNTA(tblGelirGider[Tarih])>0,SubeHedefMarj,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Toplam Nakit Pozisyonu (TL)", yazi="333333")
    h(ws, 9, 2, "=HESAP!B13", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Zararlı Kayıt Sayısı", yazi="333333")
    h(ws, 10, 2, "=HESAP!B18", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblGelirGider[Tarih])=0,"VERİ YOK",'
                 'IF(OR(COUNTIF(tblGelirGider[Tutar (TL)],"<0")>0,COUNTIF(tblNakit[Tutar (TL)],"<0")>0),"DURDUR",'
                 'IF(HESAP!B9<0,"ZARARLI",'
                 'IF(HESAP!B9>=SubeHedefMarj,"KARLI",'
                 '"İNCELE"))))',
      boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 12, 2, '=IF(B11="VERİ YOK","Gelir/gider kaydı girilmedi; karar üretilemiyor.",'
                 'IF(B11="DURDUR","Negatif tutar girişi tespit edildi; veriler düzeltilmeli.",'
                 'IF(B11="ZARARLI","Toplam kâr marjı negatif; giderler gelirin üzerinde, acil önlem gereklidir.",'
                 'IF(B11="KARLI","Toplam kâr marjı hedef marjın üzerinde; şube ağı kârlıdır.",'
                 '"Kâr marjı pozitif ancak hedef marjın altında; gider kalemleri ve nakit akışı incelenmelidir."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[12].height = 44
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["Aksiyon 1", "Aksiyon 2", "Aksiyon 3"], 15):
        h(ws, i, 1, ad, yazi="333333", kalin=True)
        h(ws, i, 2, f'=IF(B11="VERİ YOK","Şube, gelir/gider ve nakit verilerini girin; karar otomatik üretilir.",'
                    f'IF(B11="DURDUR","Negatif tutar değerlerini düzeltin.",SubeAksiyon{i-14}_ACIKLAMA))',
          yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[i].height = 40
    h(ws, 19, 1, "Karar Notu", yazi="333333")
    h(ws, 19, 2, "Bu karar destek aracıdır; şube kârlılığı fiili kayıtlarla doğrulanmalıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:B19", URUN_AD)
    alt_bant(ws, 21, "Karar kuralları KILAVUZ sayfasında gösterilmiştir.")
    genislik(ws, {"A": 32, "B": 80})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (1, "Toplam Gelir (TL)", "=HESAP!B6", TL),
        (2, "Toplam Gider (TL)", "=HESAP!B7", TL),
        (3, "TOPLAM NET KÂR (TL)", "=HESAP!B8", TL),
        (4, "Toplam Kâr Marjı", "=HESAP!B9", YÜZDE),
        (5, "Toplam Nakit Pozisyonu (TL)", "=HESAP!B13", TL),
        (6, "Nakit Oranı", "=HESAP!B22", YÜZDE),
        (7, "Karar", "=KARAR!B11", NORMAL),
        (8, "Kalite Skoru", "=KONTROLLER!B17", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En Yüksek Gider (TL)", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 6, 5, "=HESAP!B17", sayi=TL, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "Zararlı Kayıt Sayısı", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 6, 8, "=HESAP!B18", sayi=CATI, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    ws.row_dimensions[6].height = 30
    # Şube bazlı gelir/gider/kâr
    h(ws, 10, 1, "Şube", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Gelir (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Gider (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Kâr Marjı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, etiket in enumerate(SUBE_ADLARI, 11):
        h(ws, i, 1, etiket, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!F{i+16}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=IFERROR(HESAP!F{i+16}/HESAP!B{i+16},0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Şube Bazlı Gelir ve Gider"
    g1.add_data(Reference(ws, min_col=2, max_col=3, min_row=10, max_row=17), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g1.height, g1.width = 9, 15
    ws.add_chart(g1, "F8")
    g2 = BarChart()
    g2.type = "col"
    g2.title = "Şube Bazlı Net Kâr"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=17), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Şube Bazlı Kâr Marjı"
    g7.add_data(Reference(ws, min_col=5, min_row=10, max_row=17), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g7.height, g7.width = 9, 12
    ws.add_chart(g7, "AH8")
    # Gelir/gider dağılımı
    h(ws, 19, 1, "Dağılım", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 2, "Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 3, "Pay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, kalem in enumerate(["Gelir", "Gider"], 20):
        h(ws, i, 1, kalem, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i-14}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=IFERROR(HESAP!B{i-14}/(HESAP!B6+HESAP!B7),0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g3 = PieChart()
    g3.title = "Gelir / Gider Dağılımı"
    g3.add_data(Reference(ws, min_col=2, min_row=19, max_row=22), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=20, max_row=22))
    g3.height, g3.width = 9, 12
    ws.add_chart(g3, "F21")
    # Nakit trendi
    h(ws, 26, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 26, 2, "Nakit Akışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(4):
        h(ws, 27 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 27 + i, 2, f"=HESAP!B{43+i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Nakit Akışı Trendi"
    g4.add_data(Reference(ws, min_col=2, min_row=26, max_row=31), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=27, max_row=31))
    g4.height, g4.width = 9, 15
    ws.add_chart(g4, "F31")
    # Gider kalemleri
    h(ws, 33, 1, "Gider Kalemi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 33, 2, "Tutar (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, kalem in enumerate(["Kira Gideri", "Personel", "Elektrik", "Su", "Malzeme", "Diğer"], 34):
        h(ws, i, 1, kalem, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i+1}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Gider Kalemlerinin Dağılımı"
    g8.add_data(Reference(ws, min_col=2, min_row=33, max_row=40), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=1, min_row=34, max_row=40))
    g8.height, g8.width = 9, 15
    ws.add_chart(g8, "F43")
    # Projeksiyon
    h(ws, 42, 1, "Ortalama Nakit Akışı (TL)", yazi="333333")
    h(ws, 42, 2, "=HESAP!B48", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 43, 1, "Gelecek Dönem Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 43, 2, "=HESAP!B49", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 44, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 44, 2, "=HESAP!B51", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 45, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 45, 2, "=HESAP!B52", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Tahmin ve Bant"
    g5.add_data(Reference(ws, min_col=2, min_row=42, max_row=45), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=43, max_row=45))
    g5.height, g5.width = 9, 15
    ws.add_chart(g5, "U43")
    # Senaryo verisi
    h(ws, 47, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 47, 6, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 48):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-41}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = LineChart()
    g6.title = "Senaryo Net Kârları"
    g6.add_data(Reference(ws, min_col=6, min_row=47, max_row=52), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=5, min_row=48, max_row=52))
    g6.height, g6.width = 9, 12
    ws.add_chart(g6, "F55")
    baski_hazirla(ws, "A1:M55", URUN_AD)
    alt_bant(ws, 55, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 30, "B": 20, "C": 22, "D": 18, "E": 20, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Hasılat çarpanına göre senaryo bant aralığı ve gider/ciro duyarlılık "
                "(tornado) analizi.", yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Hasılat Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        (7, "İyimser", "=IF(COUNTA(tblGelirGider[Tarih])>0,SubeSenIyi,0)",
         "=IFERROR(HESAP!B8*SubeSenIyi/SubeSenBaz,0)"),
        (8, "Baz", "=IF(COUNTA(tblGelirGider[Tarih])>0,SubeSenBaz,0)", "=HESAP!B8"),
        (9, "Kötümser", "=IF(COUNTA(tblGelirGider[Tarih])>0,SubeSenKotu,0)",
         "=IFERROR(HESAP!B8*SubeSenKotu/SubeSenBaz,0)"),
        (10, "Kritik", "=IF(COUNTA(tblGelirGider[Tarih])>0,SubeSenKritik,0)",
         "=IFERROR(HESAP!B8*SubeSenKritik/SubeSenBaz,0)"),
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
        (14, "Gider Kalemleri", "=HESAP!B8*SubeTornadoGiderOran", "=HESAP!B8*SubeTornadoGiderOran*3"),
        (15, "Ciro", "=HESAP!B8*SubeTornadoCiroOran", "=HESAP!B8*SubeTornadoCiroOran*3"),
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
    h(ws, 4, 1, '=CONCATENATE("Rapor Tarihi: ",TEXT(SubeRaporTarihi,"gg.aa.yyyy")," | ",SubeFirmaUnvan)',
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Karar Gerekçesi", yazi="333333", kaydir=True)
    h(ws, 7, 2, "=KARAR!B12", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[7].height = 40
    ozet = [
        (8, "Toplam Gelir (TL)", "=HESAP!B6"),
        (9, "Toplam Gider (TL)", "=HESAP!B7"),
        (10, "Toplam Net Kâr (TL)", "=HESAP!B8"),
        (11, "Toplam Kâr Marjı", "=HESAP!B9"),
        (12, "Toplam Nakit Pozisyonu (TL)", "=HESAP!B13"),
        (13, "Nakit Oranı", "=HESAP!B22"),
        (14, "Şube Sayısı", "=HESAP!B16"),
        (15, "Zararlı Kayıt Sayısı", "=HESAP!B18"),
    ]
    for satir, ad, form in ozet:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=TL if "TL" in ad else YÜZDE if "Oran" in ad or "Marj" in ad else CATI,
          yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 18 + i, 1, f"Aksiyon {i+1}", yazi="333333", kalin=True)
        h(ws, 18 + i, 2, f"=KARAR!B{15+i}", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[18 + i].height = 38
    h(ws, 23, 1, "Hazırlayan", yazi="333333")
    h(ws, 23, 2, "=SubeRaporHazirlayan", yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Sürüm", yazi="333333")
    h(ws, 24, 2, "=SubeDosyaSurumu", yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Bu rapor karar destek amaçlıdır; nihai karar fiili kayıtlarla doğrulanmalıdır.",
      yazi=GRİ, boyut=9, kaydir=True)
    baski_hazirla(ws, "A1:B26", URUN_AD)
    genislik(ws, {"A": 40, "B": 26})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=36)
    h(ws, 3, 1, "Örnek Veriler", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Ürünü tanımak için bu sayfadaki örnek şube, gelir/gider ve nakit verileri "
                "yan yana bloklar halinde sunulur; ana sayfalarda denemeler yapın.",
      yazi=GRİ, kaydir=True)
    # ---- Blok 1: Örnek Şubeler (A-F) ----
    h(ws, 5, 1, "Örnek Şubeler", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar = ["Şube Adı", "Şehir", "Açılış Nakit (TL)", "Hedef Kâr Marjı", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar, basla=1)
    formuller = {
        "Açıklama": '=IF(tblOrnekSube[[#This Row],[Şube Adı]]="","",CONCATENATE("Şube: ",tblOrnekSube[[#This Row],[Şube Adı]]," | Şehir: ",tblOrnekSube[[#This Row],[Şehir]]))',
    }
    tablo_ekle(ws, "tblOrnekSube", "A6:E1005", sutunlar, formuller)
    ornek_subeler = [
        ("İstanbul Merkez", "İstanbul", 250000, 0.15),
        ("Ankara Şubesi", "Ankara", 180000, 0.12),
        ("İzmir Şubesi", "İzmir", 220000, 0.13),
        ("Bursa Şubesi", "Bursa", 120000, 0.10),
        ("Antalya Şubesi", "Antalya", 160000, 0.14),
        ("Adana Şubesi", "Adana", 110000, 0.11),
    ]
    for i, (ad, sehir, nakit, marj) in enumerate(ornek_subeler, 7):
        for kolon, deger in [(1, ad), (2, sehir), (3, nakit), (4, marj)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = YÜZDE
    giris_hucreleri(ws, 7, 1005, [1, 2, 3, 4])
    # ---- Blok 2: Örnek Gelir/Gider (H-N) ----
    h(ws, 5, 8, "Örnek Gelir ve Giderler", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar2 = ["Tarih", "Şube", "Kalem", "Tür", "Tutar (TL)", "İşaretli Tutar (TL)", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar2, basla=8)
    formuller2 = {
        "İşaretli Tutar (TL)": '=IF(tblOrnekGelirGider[[#This Row],[Tutar (TL)]]="","",IF(tblOrnekGelirGider[[#This Row],[Tür]]="Gider",-tblOrnekGelirGider[[#This Row],[Tutar (TL)]],tblOrnekGelirGider[[#This Row],[Tutar (TL)]]))',
        "Açıklama": '=IF(tblOrnekGelirGider[[#This Row],[Tarih]]="","",CONCATENATE("Kayıt: ",TEXT(tblOrnekGelirGider[[#This Row],[Tarih]],"gg.aa.yyyy")," | ",tblOrnekGelirGider[[#This Row],[Kalem]]," | ",tblOrnekGelirGider[[#This Row],[Tür]]))',
    }
    tablo_ekle(ws, "tblOrnekGelirGider", "H6:N1005", sutunlar2, formuller2)
    ornek_kayit = [
        (date(2026, 6, 5), "İstanbul Merkez", "Satış", "Gelir", 480000),
        (date(2026, 6, 5), "İstanbul Merkez", "Kira Gideri", "Gider", 45000),
        (date(2026, 6, 8), "Ankara Şubesi", "Satış", "Gelir", 310000),
        (date(2026, 6, 8), "Ankara Şubesi", "Personel", "Gider", 120000),
        (date(2026, 6, 10), "İzmir Şubesi", "Hizmet", "Gelir", 265000),
        (date(2026, 6, 10), "İzmir Şubesi", "Elektrik", "Gider", 22000),
        (date(2026, 6, 12), "Bursa Şubesi", "Satış", "Gelir", 185000),
        (date(2026, 6, 12), "Bursa Şubesi", "Malzeme", "Gider", 75000),
        (date(2026, 6, 15), "Antalya Şubesi", "Satış", "Gelir", 240000),
        (date(2026, 6, 15), "Antalya Şubesi", "Personel", "Gider", 90000),
        (date(2026, 6, 18), "Adana Şubesi", "Hizmet", "Gelir", 155000),
        (date(2026, 6, 18), "Adana Şubesi", "Su", "Gider", 18000),
    ]
    for i, (tarih, sube, kalem, tur, tutar) in enumerate(ornek_kayit, 7):
        for kolon, deger in [(8, tarih), (9, sube), (10, kalem), (11, tur), (12, tutar)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=8).number_format = TARİH
        ws.cell(row=satir, column=12).number_format = TL
        ws.cell(row=satir, column=13).number_format = TL
    giris_hucreleri(ws, 7, 1005, [8, 9, 10, 11, 12])
    # ---- Blok 3: Örnek Nakit (P-U) ----
    h(ws, 5, 16, "Örnek Nakit Akışı", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar3 = ["Tarih", "Şube", "İşlem Türü", "Tutar (TL)", "İşaretli Tutar (TL)", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar3, basla=16)
    formuller3 = {
        "İşaretli Tutar (TL)": '=IF(tblOrnekNakit[[#This Row],[Tutar (TL)]]="","",IF(tblOrnekNakit[[#This Row],[İşlem Türü]]="Ödeme",-tblOrnekNakit[[#This Row],[Tutar (TL)]],tblOrnekNakit[[#This Row],[Tutar (TL)]]))',
        "Açıklama": '=IF(tblOrnekNakit[[#This Row],[Tarih]]="","",CONCATENATE("İşlem: ",TEXT(tblOrnekNakit[[#This Row],[Tarih]],"gg.aa.yyyy")," | ",tblOrnekNakit[[#This Row],[İşlem Türü]]," | Tutar: ",TEXT(tblOrnekNakit[[#This Row],[Tutar (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblOrnekNakit", "P6:U1005", sutunlar3, formuller3)
    ornek_nakit = [
        (date(2026, 6, 5), "İstanbul Merkez", "Tahsilat", 480000),
        (date(2026, 6, 6), "İstanbul Merkez", "Ödeme", 45000),
        (date(2026, 6, 8), "Ankara Şubesi", "Tahsilat", 310000),
        (date(2026, 6, 9), "Ankara Şubesi", "Ödeme", 120000),
        (date(2026, 6, 10), "İzmir Şubesi", "Tahsilat", 265000),
        (date(2026, 6, 11), "İzmir Şubesi", "Ödeme", 22000),
        (date(2026, 6, 12), "Bursa Şubesi", "Tahsilat", 185000),
        (date(2026, 6, 13), "Bursa Şubesi", "Ödeme", 75000),
        (date(2026, 6, 15), "Antalya Şubesi", "Tahsilat", 240000),
        (date(2026, 6, 16), "Antalya Şubesi", "Ödeme", 90000),
        (date(2026, 6, 18), "Adana Şubesi", "Tahsilat", 155000),
        (date(2026, 6, 19), "Adana Şubesi", "Ödeme", 18000),
    ]
    for i, (tarih, sube, tur, tutar) in enumerate(ornek_nakit, 7):
        for kolon, deger in [(16, tarih), (17, sube), (18, tur), (19, tutar)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=16).number_format = TARİH
        ws.cell(row=satir, column=19).number_format = TL
        ws.cell(row=satir, column=20).number_format = TL
    giris_hucreleri(ws, 7, 1005, [16, 17, 18, 19])
    sabitle(ws, "A7")
    genislik(ws, {"A": 20, "B": 14, "C": 18, "D": 16, "E": 30,
                  "H": 12, "I": 20, "J": 14, "K": 10, "L": 13, "M": 17, "N": 40,
                  "P": 12, "Q": 20, "R": 12, "S": 13, "T": 17, "U": 40})


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
    h(ws, 6, 1, "Şubeler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Şehirler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 5, "Kalemler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 7, "Türler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 9, "İşlem Türleri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ad in enumerate(SUBE_ADLARI, 7):
        h(ws, i, 1, ad, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Şube adı | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni şube için satır ekleyin")
    for i, sehir in enumerate(SEHIRLER, 7):
        h(ws, i, 3, sehir, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"C{i}", "Tanım: Şehir | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni şehir için satır ekleyin")
    for i, kalem in enumerate(KALEMLER, 7):
        h(ws, i, 5, kalem, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"E{i}", "Tanım: Kalem | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni kalem için satır ekleyin")
    for i, tur in enumerate(["Gelir", "Gider"], 7):
        h(ws, i, 7, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"G{i}", "Tanım: Kayıt türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Gelir veya Gider girin")
    for i, tur in enumerate(["Tahsilat", "Ödeme"], 7):
        h(ws, i, 9, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"I{i}", "Tanım: İşlem türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Tahsilat veya Ödeme girin")
    genislik(ws, {"A": 22, "C": 16, "E": 18, "G": 12, "I": 14})


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
        (7, "SubeFirmaUnvan", "Örnek Perakende A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "SubeRaporTarihi", TARIH_D, "Tarih", None, None, "Raporun dayanak tarihi.", "Kullanıcı"),
        (9, "SubeRaporHazirlayan", "Mali İşler ve Operasyon Yöneticisi", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "SubeDosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "SubeHedefMarj", 0.15, "Oran", 0, 1, "Hedef kâr marjı eşiği (%15).", "Varsayım"),
        (12, "SubeKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (13, "SubeKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (14, "SubeSenIyi", 0.90, "Oran", 0, 3, "İyimser senaryo hasılat çarpanı.", "Varsayım"),
        (15, "SubeSenBaz", 1.00, "Oran", 0, 3, "Baz senaryo hasılat çarpanı.", "Varsayım"),
        (16, "SubeSenKotu", 1.15, "Oran", 0, 3, "Kötümser senaryo hasılat çarpanı.", "Varsayım"),
        (17, "SubeSenKritik", 1.30, "Oran", 0, 3, "Kritik senaryo hasılat çarpanı.", "Varsayım"),
        (18, "SubeTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (19, "SubeTahminDonem", 5, "Dönem", 1, 12, "Gelecek dönem tahmin noktası.", "Varsayım"),
        (20, "SubeTornadoGiderOran", 0.10, "Oran", 0, 0.5, "Gider kalemleri tornado adımı (%10).", "Varsayım"),
        (21, "SubeTornadoCiroOran", 0.05, "Oran", 0, 0.5, "Ciro tornado adımı.", "Varsayım"),
        (22, "SubeP90Oran", 0.90, "Oran", 0, 1, "Kayıt tutarı dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (23, "SubeYuksekGiderEsik", 100000, "TL", 0, 1000000000, "Yüksek gider tutarı kontrol eşiği.", "Varsayım"),
        (24, "SubeYeterliVeri", 25, "Adet", 1, 100000, "Veri kalite skorunun %100 sayılması için gerekli kayıt sayısı.", "Varsayım"),
        (25, "SubeAksiyon1_ACIKLAMA",
         "Kâr marjı negatif; zarar eden şubelerin gelir/gider yapısını kalem kalem denetleyin, düşük ciro ve yüksek sabit gider kalemlerini yeniden yapılandırın.", "Metin", None, None,
         "ZARARLI kararında gösterilen ilk öneri.", "Varsayım"),
        (26, "SubeAksiyon2_ACIKLAMA",
         "Kâr marjı hedefin altında; gider kalemleri ve nakit ödeme planlamasını gözden geçirin, tahsilat süresini kısaltacak önlemleri uygulayın.", "Metin", None, None,
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (27, "SubeAksiyon3_ACIKLAMA",
         "Kâr marjı hedefin üzerinde; mevcut şube performansını sürdürerek dönemlik nakit ve kârlılık eğilimini izlemeye devam edin.", "Metin", None, None,
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
        h(ws, satir, 1, f"SubePanelSatir_{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki grafiğin şube sırası ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel sıra numarası {i} | "
                   f"Neden önemli: Grafikleri besler | "
                   f"Doğru kullanım: Şube sırasını girin")
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
        "Gelir/gider kaydı girilmemişse karar VERİ YOK olur; negatif tutar varsa DURDUR olur.",
    ]
    for i, m in enumerate(kurallar, 6):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=28)
    h(ws, 12, 1, "Hesaplama Yöntemleri", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    yontemler = [
        "İşaretli tutar = Gelir için +Tutar, Gider için −Tutar.",
        "Net kâr = Toplam gelir − Toplam gider; kâr marjı = Net kâr / Toplam gelir.",
        "Nakit akışı = Toplam tahsilat − Toplam ödeme.",
        "Nakit pozisyonu = Açılış nakitleri toplamı + Nakit akışı.",
        "En yüksek gider = Gider türündeki en büyük kayıt tutarı (MAXIFS).",
        "Gelecek dönem nakit tahmini = FORECAST.LINEAR ile trend uzatması.",
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
    h(ws, 27, 1, "Bu dosya karar destek aracıdır; şube kârlılığı ve nakit durumu fiili "
                 "kayıtlarla doğrulanmalıdır.", yazi=GRİ, boyut=9, kaydir=True)
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
    for satir in range(6, 13):
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil))
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"DOLU"'], fill=yesil))
        ws3.conditional_formatting.add(f"B{satir}",
            CellIsRule(operator="equal", formula=['"BOŞ"'], fill=sari))
    ws3.conditional_formatting.add("B17",
        CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil))
    ws3.conditional_formatting.add("B17",
        CellIsRule(operator="between", formula=["70", "89"], fill=sari))
    ws3.conditional_formatting.add("B17",
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
    for satir in range(48, 52):
        ws5.conditional_formatting.add(f"F{satir}",
            CellIsRule(operator="greaterThan", formula=["0"], fill=yesil))
        ws5.conditional_formatting.add(f"F{satir}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))


def moduller(wb):
    ws = wb["KONTROLLER"]
    modul_satirlari = {
        "modulT1OrtMarj": ("T1", "Ortalama Kâr Marjı",
                           "=HESAP!B9",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama kâr marjı ",TEXT(C24,"0.0%")," — şube ağının genel kârlılık düzeyini gösterir")',
                           YÜZDE),
        "modulT2OrtNakit": ("T2", "Ortalama Nakit Akışı",
                            "=HESAP!B12",
                            '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama nakit akışı ",TEXT(C25,"₺ #,##0")," TL — tahsilat ve ödeme dengesini gösterir")',
                            TL),
        "modulO1GiderSapmasi": ("O1", "Gider Sapması (MAD)",
                                "=IFERROR(AVEDEV(tblGelirGider[Tutar (TL)]),0)",
                                '=_xlfn.TEXTJOIN("; ",TRUE,"Kayıt tutarları ortalama mutlak sapması ",TEXT(C26,"₺ #,##0")," TL — kayıtlar arası heterojenliği ölçer")',
                                TL),
        "modulO2EnYuksekGider": ("O2", "En Yüksek Gider",
                                 "=HESAP!B17",
                                 '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek gider kaydı ",TEXT(C27,"₺ #,##0")," TL — öncelikli kontrol kalemidir")',
                                 TL),
        "modulO3Senaryo": ("O3", "Senaryo Bant Genişliği",
                           "=IFERROR(STDEV.P(SENARYO_DUYARLILIK!C7:C10),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo net kâr standart sapması ",TEXT(C28,"₺ #,##0")," TL — kâr belirsizliğinin etkisini gösterir")',
                           TL),
        "modulO6Hhi": ("O6", "Gelir Yoğunlaşması (HHI)",
                       '=IFERROR(IF(SUM(HESAP!B27:B32)=0,0,SUMPRODUCT((HESAP!B27:B32/SUM(HESAP!B27:B32))^2)),0)',
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Şube gelir yoğunlaşma endeksi ",TEXT(C29,"0.000")," — 1 e yaklaştıkça gelir tek şubede toplanır")',
                       None),
        "modulO8KaliteSkor": ("O8", "Veri Kalite Skoru",
                              "=KONTROLLER!B17",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C30,"0")," puan — ",IF(C30>=SubeKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))',
                              CATI),
        "modulI1Tahmin": ("I1", "Gelecek Dönem Nakit Tahmini",
                          "=HESAP!B49",
                          '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem nakit tahmini ",TEXT(C31,"₺ #,##0")," TL — bant aralığı için alt/üst sınır HESAP sayfasındadır")',
                          TL),
        "modulI2P90": ("I2", "Kayıt Tutarı Yüzdelik P90",
                       "=IFERROR(_xlfn.PERCENTILE.INC(tblGelirGider[Tutar (TL)],SubeP90Oran),0)",
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Kayıt tutarlarının %90 ı ",TEXT(C32,"₺ #,##0")," TL altındadır — tutar dağılımının üst sınırını gösterir")',
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
    sayfa_sirasi = ["KAPAK", "HIZLI_BASLANGIC", "SUBE_LISTESI", "GELIR_GIDER", "NAKIT_AKISI",
                    "HESAP", "KONTROLLER", "KARAR", "PANO", "SENARYO_DUYARLILIK", "RAPOR",
                    "ORNEK_VERI", "DEGISIKLIK_KAYDI", "LISTELER", "AYARLAR", "KILAVUZ"]
    for ad in sayfa_sirasi:
        wb.create_sheet(ad)
    ilk = wb["Sheet"]
    wb.remove(ilk)

    kapak(wb["KAPAK"])
    hizli_baslangic(wb["HIZLI_BASLANGIC"])
    sube_listesi(wb["SUBE_LISTESI"])
    gelir_gider(wb["GELIR_GIDER"])
    nakit_akisi(wb["NAKIT_AKISI"])
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
        if isinstance(anahtar, str) and anahtar.startswith("Sube"):
            ad_ekle(wb, anahtar, f"AYARLAR!$B${satir}")

    # Kritik çıktı ad tanımları (D02 karşılıkları ve PANO beslemeleri)
    ad_ekle(wb, "SubeToplamGelir", "HESAP!$B$6")
    ad_ekle(wb, "SubeToplamGider", "HESAP!$B$7")
    ad_ekle(wb, "SubeNetKar", "HESAP!$B$8")
    ad_ekle(wb, "SubeNakitPozisyonu", "HESAP!$B$13")
    ad_ekle(wb, "SubeKarar", "KARAR!$B$11")
    ad_ekle(wb, "SubeProjX", "HESAP!$A$43:$A$46")

    ad_ekle(wb, "ListeSubeler", "LISTELER!$A$7:$A$12")
    ad_ekle(wb, "ListeSehirler", "LISTELER!$C$7:$C$14")
    ad_ekle(wb, "ListeKalemler", "LISTELER!$E$7:$E$14")
    ad_ekle(wb, "ListeTurler", "LISTELER!$G$7:$G$8")
    ad_ekle(wb, "ListeIslemTurleri", "LISTELER!$I$7:$I$8")

    for ws in wb.worksheets:
        sayfa_koru(ws)

    dosya = "SubeKarlilikVeNakitHesaplayici.xlsx"
    wb.save(dosya)
    print("Dosya oluşturuldu:", dosya)


if __name__ == "__main__":
    main()
