"""
Vergi/SGK Borcunu Tecil Etmeli Miyim? (Kredi mi Tecil mi?) — üretim betiği.
Manda v4 uyumlu: sayfa sırası, sarı giriş, 1234 koruma, ipucu/not, tablo mimarisi.
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
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

URUN_AD = "Vergi/SGK Borcunu Tecil Etmeli Miyim? (Kredi mi Tecil mi?)"
SURUM = "1.0.0"
RENK = "1F7A4D"

BORC_TURLERI = ["Gelir Vergisi", "Kurumlar Vergisi", "KDV", "Muhtasar Stopaj",
                "SGK Primi", "SGK Cezası", "Diğer"]
KARAR_SECENEKLERI = ["TECİL YAP", "KREDİ KULLAN", "PEŞİN ÖDE"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Vergi ve SGK borçlarınızı girin; tecil, banka kredisi ve peşin ödeme "
                "seçeneklerinin toplam maliyetini ve bugünkü değerini karşılaştırın; "
                "size en uygun ödeme yoluna karar verin.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Borç aslı, gecikme zammı ve vade yapısını tek tabloda izler; toplam borç ve aylık zam yükünü hesaplar",
        "Tecil (vadeye yayma), banka kredisi ve peşin ödeme seçeneklerinin toplam maliyetini hesaplar",
        "Seçenekleri bugünkü değerine (NBD) indirgeyerek en ucuz ödeme yolunu belirler",
        "Tecil faizi, kredi faizi ve iskonto oranına göre senaryo bant aralığı ve tornado analizi",
        "Borç türü yoğunlaşması, anomali ve veri kalite skoru ile karar kapısı ve aksiyon önerisi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Vergi ve SGK borcu için tecil başvurusu değerlendiren mükellefler",
        "Borcunu kredi ile kapatmayı düşünen işletme sahipleri ve finans yöneticileri",
        "Tecil ile kredi arasında maliyet karşılaştırması yapmak isteyen mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. BORC_LISTESI sayfasına borçlarınızı girin. "
                 "3. PANO ve RAPOR sayfalarından önerilen ödeme yolunu izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Borçlarınızı girin",
         "BORC_LISTESI sayfasındaki sarı hücrelere borç türü, borç aslı, gecikme zammı, vade "
         "tarihi ve aylık zam oranını girin."),
        ("Adım 2 — Seçenekleri izleyin",
         "HESAP sayfası tecil, kredi ve peşin ödeme seçeneklerinin toplam maliyetini ve NBD'sini hesaplar."),
        ("Adım 3 — Kararı değerlendirin",
         "KARAR sayfası en düşük maliyetli ödeme yolunu gerekçesiyle önerir."),
        ("Adım 4 — Oranları ayarlayın",
         "Tecil faizi, kredi faizi, iskonto oranı ve terkin oranı AYARLAR sayfasından değiştirilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Gecikme zammını borç aslının dışında eksik girmek (toplam borç eksik çıkar)",
        "Vade tarihini rapor tarihinden sonra bırakmak (geciken ay sıfır görünür)",
        "Aylık zam oranını yüzde yerine 1'in üzerinde ondalık girmek (zam yükü şişer)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; tecil başvurusu sonucu bağlayıcı değildir.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def borc_listesi(ws):
    sayfa_hazirla(ws, "BORC_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Vergi ve SGK Borç Listesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Her satır bir vergi/SGK borç kalemidir; "
                "geciken ay, toplam borç ve aylık zam yükü otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Borç Kodu", "Borç Türü", "Borç Aslı (TL)", "Gecikme Zammı (TL)",
                "Vade Tarihi", "Aylık Zam Oranı", "Geciken Ay", "Toplam Borç (TL)",
                "Aylık Zam Yükü (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Geciken Ay": "=IF(tblBorc[[#This Row],[Vade Tarihi]]=\"\",0,"\
                      "MAX(0,ROUNDDOWN((TecilRaporTarihi-tblBorc[[#This Row],[Vade Tarihi]])/30,0)))",
        "Toplam Borç (TL)": "=IF(tblBorc[[#This Row],[Borç Aslı (TL)]]=\"\",0,"\
                            "tblBorc[[#This Row],[Borç Aslı (TL)]]+tblBorc[[#This Row],[Gecikme Zammı (TL)]])",
        "Aylık Zam Yükü (TL)": "=IF(tblBorc[[#This Row],[Borç Aslı (TL)]]=\"\",0,"\
                               "tblBorc[[#This Row],[Borç Aslı (TL)]]*tblBorc[[#This Row],[Aylık Zam Oranı]])",
        "Açıklama": "=IF(tblBorc[[#This Row],[Borç Türü]]=\"\",\"\","\
                    "CONCATENATE(\"Borç: \",tblBorc[[#This Row],[Borç Türü]],\" | Toplam: \",TEXT(tblBorc[[#This Row],[Toplam Borç (TL)]],\"#,##0\")))",
    }
    tablo_ekle(ws, "tblBorc", "A5:J1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeBorcTurleri", "B6:B1004",
              baslik="Borç Türü", mesaj="Listeden bir borç türü seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Borç Aslı", mesaj="Borç aslını TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Gecikme Zammı", mesaj="Bugüne kadar işlemiş gecikme zammını TL girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "date", "01.01.2020", "E6:E1004",
              baslik="Vade Tarihi", mesaj="Borç kaleminin vade tarihini girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2020 ile 31.12.2100 arasında olmalı.",
              isaret="between", f2="31.12.2100")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Aylık Zam Oranı", mesaj="Aylık gecikme zammı oranını 0-1 arasında girin (örn. %4 için 0,04).",
              hata_baslik="Geçersiz oran", hata_mesaj="Oran 0 ile 1 arasında olmalı.",
              isaret="between", f2="1")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=5).number_format = TARİH
        ws.cell(row=satir, column=6).number_format = YÜZDE
    ornek_borclar = [
        ("V-1001", "KDV", 150000, 42000, date(2025, 3, 26), 0.04),
        ("V-1002", "Muhtasar Stopaj", 80000, 19000, date(2025, 6, 26), 0.04),
        ("V-1003", "Kurumlar Vergisi", 240000, 61000, date(2025, 5, 31), 0.04),
        ("S-2001", "SGK Primi", 120000, 31000, date(2025, 8, 31), 0.03),
        ("S-2002", "SGK Cezası", 40000, 14000, date(2025, 10, 31), 0.04),
        ("V-1004", "Gelir Vergisi", 60000, 9000, date(2025, 11, 30), 0.04),
        ("V-1005", "KDV", 90000, 8000, date(2026, 2, 26), 0.04),
    ]
    giris_kolonlari = [1, 2, 3, 4, 5, 6]
    for i, satir in enumerate(ornek_borclar, 6):
        for k, deger in enumerate(satir):
            ws.cell(row=i, column=giris_kolonlari[k]).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 26, "B": 20, "C": 18, "D": 18, "E": 13, "F": 14, "G": 12,
                  "H": 18, "I": 18, "J": 30})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Borç Türü Bazlı Özet ve Ödeme Seçeneği Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet borç listesinden otomatik beslenir; seçenek hesapları "
                "AYARLAR'daki oranlarla canlı üretilir.", yazi=GRİ, kaydir=True)
    sutunlar = ["Borç Türü", "Asıl Toplam", "Zam Toplam", "Toplam Borç",
                "Aylık Zam Yükü", "Ortalama Geciken Ay"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(BORC_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=SUMIFS(tblBorc[Borç Aslı (TL)],tblBorc[Borç Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblBorc[Gecikme Zammı (TL)],tblBorc[Borç Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblBorc[Toplam Borç (TL)],tblBorc[Borç Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblBorc[Aylık Zam Yükü (TL)],tblBorc[Borç Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=IFERROR(AVERAGEIFS(tblBorc[Geciken Ay],tblBorc[Borç Türü],$A{r}),0)',
          sayi="0.0", yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "TOPLAM BLOK", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Toplam Borç Aslı (TL)", yazi="333333")
    h(ws, 15, 2, "=SUM(B6:B12)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 16, 1, "Toplam Gecikme Zammı (TL)", yazi="333333")
    h(ws, 16, 2, "=SUM(C6:C12)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "TOPLAM BORÇ (TL)", yazi="333333", kalin=True)
    h(ws, 17, 2, "=SUM(D6:D12)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Toplam Aylık Zam Yükü (TL)", yazi="333333")
    h(ws, 18, 2, "=SUM(E6:E12)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Ortalama Geciken Ay", yazi="333333")
    h(ws, 19, 2, "=IFERROR(AVERAGE(tblBorc[Geciken Ay]),0)", sayi="0.0", yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Borç Kalemi Sayısı", yazi="333333")
    h(ws, 20, 2, "=COUNTA(tblBorc[Borç Kodu])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 21, 1, "Borç Türü Sayısı (farklı)", yazi="333333")
    h(ws, 21, 2, "=IFERROR(SUMPRODUCT((tblBorc[Borç Türü]<>\"\")/COUNTIF(tblBorc[Borç Türü],tblBorc[Borç Türü]&\"\")),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    # Tecil seçeneği
    h(ws, 23, 1, "TECİL SEÇENEĞİ (6183/48 kapsamında vadeye yayma)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A23:F23")
    h(ws, 24, 1, "Peşinat Oranı", yazi="333333")
    h(ws, 24, 2, "=AYARLAR!$B$11", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Peşinat Tutarı (TL)", yazi="333333")
    h(ws, 25, 2, "=B17*B24", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Tecil Anaparası (TL)", yazi="333333")
    h(ws, 26, 2, "=B17-B25", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Aylık Tecil Faizi", yazi="333333")
    h(ws, 27, 2, "=AYARLAR!$B$13/12", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Tecil Taksit Sayısı", yazi="333333")
    h(ws, 28, 2, "=AYARLAR!$B$12", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 29, 1, "Aylık Taksit (TL)", yazi="333333", kalin=True)
    h(ws, 29, 2, "=IF(OR(B26<=0,B28=0),0,B26*B27/(1-(1+B27)^-B28))", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 30, 1, "Tecil Toplam Maliyet (TL)", yazi="333333", kalin=True)
    h(ws, 30, 2, "=B29*B28+B25", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    # Kredi seçeneği
    h(ws, 32, 1, "KREDİ SEÇENEĞİ (bankadan borcu kapatacak kredi)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A32:F32")
    h(ws, 33, 1, "Kredi Komisyon Oranı", yazi="333333")
    h(ws, 33, 2, "=AYARLAR!$B$15", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 34, 1, "Komisyon Tutarı (TL)", yazi="333333")
    h(ws, 34, 2, "=B17*B33", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 35, 1, "Kredi Anaparası (TL)", yazi="333333")
    h(ws, 35, 2, "=B17", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 36, 1, "Aylık Kredi Faizi", yazi="333333")
    h(ws, 36, 2, "=AYARLAR!$B$14/12", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 37, 1, "Kredi Taksit Sayısı", yazi="333333")
    h(ws, 37, 2, "=AYARLAR!$B$16", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 38, 1, "Aylık Taksit (TL)", yazi="333333", kalin=True)
    h(ws, 38, 2, "=IF(OR(B35<=0,B37=0),0,B35*B36/(1-(1+B36)^-B37))", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 39, 1, "Kredi Toplam Maliyet (TL)", yazi="333333", kalin=True)
    h(ws, 39, 2, "=B38*B37+B34", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    # Peşin seçeneği
    h(ws, 41, 1, "PEŞİN ÖDEME SEÇENEĞİ (yapılandırma terkinli)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A41:F41")
    h(ws, 42, 1, "Gecikme Zammı Terkin Oranı", yazi="333333")
    h(ws, 42, 2, "=AYARLAR!$B$18", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 43, 1, "Terkin Tutarı (TL)", yazi="333333")
    h(ws, 43, 2, "=B16*B42", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 44, 1, "Peşin Ödeme (TL)", yazi="333333", kalin=True)
    h(ws, 44, 2, "=B17-B43", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    # Karşılaştırma
    h(ws, 46, 1, "SEÇENEK KARŞILAŞTIRMASI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A46:F46")
    h(ws, 47, 1, "Tecil Toplam Maliyet (TL)", yazi="333333")
    h(ws, 47, 2, "=B30", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 48, 1, "Kredi Toplam Maliyet (TL)", yazi="333333")
    h(ws, 48, 2, "=B39", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 49, 1, "Peşin Ödeme Tutarı (TL)", yazi="333333")
    h(ws, 49, 2, "=B44", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 50, 1, "En Düşük Maliyetli Seçenek", yazi="333333", kalin=True)
    h(ws, 50, 2, "=IF(AND(B47<=B48,B47<=B49),\"TECİL\",IF(B49<=B48,\"PEŞİN\",\"KREDİ\"))",
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 51, 1, "En Düşük Maliyet (TL)", yazi="333333")
    h(ws, 51, 2, "=MIN(B47:B49)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 52, 1, "Tecil − Kredi Farkı (TL)", yazi="333333")
    h(ws, 52, 2, "=B47-B48", sayi=TL, yazi=DIS_REF_YEŞIL)
    # NBD bloğu
    h(ws, 54, 1, "BUGÜNKÜ DEĞER KARŞILAŞTIRMASI (NBD)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A54:F54")
    h(ws, 55, 1, "NBD Tecil (TL)", yazi="333333")
    h(ws, 55, 2, '=IF(OR(B28=0,AYARLAR!$B$19/12<=0),B30,B25+B29*(1-(1+AYARLAR!$B$19/12)^-B28)/(AYARLAR!$B$19/12))',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 56, 1, "NBD Kredi (TL)", yazi="333333")
    h(ws, 56, 2, '=IF(OR(B37=0,AYARLAR!$B$19/12<=0),B39,B34+B38*(1-(1+AYARLAR!$B$19/12)^-B37)/(AYARLAR!$B$19/12))',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 57, 1, "NBD Peşin (TL)", yazi="333333")
    h(ws, 57, 2, "=B49", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 58, 1, "NBD Farkı (Tecil − Kredi) (TL)", yazi="333333", kalin=True)
    h(ws, 58, 2, "=B55-B56", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 60, "Hesap satırları yalnızca borç listesinden ve AYARLAR'dan beslenir; korumalıdır.")
    genislik(ws, {"A": 42, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Ortalama Gecikme Zammı Oranı",
         "=IFERROR(AVERAGE(tblBorc[Gecikme Zammı (TL)])/AVERAGE(tblBorc[Borç Aslı (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama gecikme zammı oranı ",TEXT(C6,"0.0%")," — zammın borç aslına oranı vade yapısının ne kadar esnediğini gösterir")'),
        ("T2", "Ortalama Toplam Borç",
         "=IFERROR(AVERAGE(tblBorc[Toplam Borç (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama toplam borç kalemi ",TEXT(C7,"₺ #,##0")," TL — borç portföyünün ortalama büyüklüğünü gösterir")'),
        ("O1", "Toplam Borç Sapması (MAD)",
         "=IFERROR(AVEDEV(tblBorc[Toplam Borç (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Toplam borç ortalama mutlak sapması ",TEXT(C8,"₺ #,##0")," TL — sapma yüksekse borç kalemleri heterojendir")'),
        ("O2", "En Yüksek Borç Kalemi",
         "=IFERROR(MAX(tblBorc[Toplam Borç (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek toplam borç kalemi ",TEXT(C9,"₺ #,##0")," TL — bu kalem tecil/kredi kararında önceliklidir")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C11",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kritik senaryo tecil maliyet farkı ",TEXT(C10,"₺ #,##0")," TL — faiz belirsizliğinin etkisini gösterir")'),
        ("O6", "Borç Yoğunlaşması (HHI)",
         "=IFERROR(IF(SUM(HESAP!$D$6:$D$12)=0,0,SUMPRODUCT((HESAP!$D$6:$D$12/SUM(HESAP!$D$6:$D$12))^2)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Borç yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça borç tek türde toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=TecilKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Borç Tahmini",
         "=PANO!B30",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem toplam borç tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Toplam Borç Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblBorc[Toplam Borç (TL)],TecilP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Toplam borç kalemlerinin %90 ı ",TEXT(C14,"₺ #,##0")," TL altında gerçekleşir — üst sınır planlamada kullanılabilir")'),
        ("I3", "NBD Farkı (Tecil − Kredi)",
         "=HESAP!B58",
         '=_xlfn.TEXTJOIN("; ",TRUE,"NBD farkı ",TEXT(C15,"₺ #,##0")," TL — fark pozitifse tecil bugünkü değerde daha pahalı, kredi avantajlıdır")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        satir = 6 + i
        h(ws, satir, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, satir, 2, ad, kaydir=True)
        h(ws, satir, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[satir].height = 30
    genislik(ws, {"A": 6, "B": 34, "C": 55, "D": 72})
    ws.sheet_view.showGridLines = True


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Giriş Kalite Kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki kontroller canlı formüllerle çalışır; sorun varsa uyarı üretir.",
      yazi=GRİ, kaydir=True)
    kontroller = [
        ("Boş borç aslı satırı var mı?", "=IF(COUNTBLANK(tblBorc[Borç Aslı (TL)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, "B3261E"),
        ("Negatif borç aslı var mı?", "=IF(COUNTIF(tblBorc[Borç Aslı (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Negatif gecikme zammı var mı?", "=IF(COUNTIF(tblBorc[Gecikme Zammı (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Vadesi 36 aydan eski borç var mı?", "=IF(COUNTIF(tblBorc[Geciken Ay],\">36\")=0,\"TEMİZ\",\"GEÇİKMİŞ\")", NORMAL, "B3261E"),
        ("Aylık zam oranı 1'in üzerinde mi?", "=IF(COUNTIF(tblBorc[Aylık Zam Oranı],\">1\")=0,\"TEMİZ\",\"ORAN >1 VAR\")", NORMAL, "B3261E"),
        ("Toplam borç eşiği aşılıyor mu?", "=IF(HESAP!B17>TecilYuksekBorcEsik,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("Tecil taksiti aylık nakit akışını aşıyor mu?", "=IF(HESAP!B29>AYARLAR!$B$27,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("NBD farkı hangi tarafı işaret ediyor?", "=IF(HESAP!B58<0,\"KREDİ AVANTAJLI\",\"TECİL AVANTAJLI\")", NORMAL, "B3261E"),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Temel giriş kolonlarının dolu olma oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblBorc[Borç Türü])*TecilTemelKolonSayisi", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblBorc[Borç Türü])+COUNTA(tblBorc[Borç Aslı (TL)])+"
                 "COUNTA(tblBorc[Gecikme Zammı (TL)])+COUNTA(tblBorc[Vade Tarihi])+"
                 "COUNTA(tblBorc[Aylık Zam Oranı])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(TecilToplamGiris=0,0,TecilDoluGiris/TecilToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=TecilKaliteIyi,\"YÜKSEK\",IF(B19>=TecilKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 60})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Tecil mi, Kredi mi, Peşin mi?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Toplam maliyet, bugünkü değer farkı ve kredi erişimine göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Toplam Borç (TL)", yazi="333333")
    h(ws, 6, 2, "=HESAP!B17", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "En Düşük Maliyetli Seçenek", yazi="333333")
    h(ws, 7, 2, "=HESAP!B50", yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "En Düşük Maliyet (TL)", yazi="333333")
    h(ws, 8, 2, "=HESAP!B51", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "NBD Farkı (Tecil − Kredi) (TL)", yazi="333333")
    h(ws, 9, 2, "=HESAP!B58", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblBorc[Borç Aslı (TL)])=0,"VERİ YOK",'\
                 'IF(OR(COUNTIF(tblBorc[Borç Aslı (TL)],"<0")>0,COUNTIF(tblBorc[Gecikme Zammı (TL)],"<0")>0),"DURDUR",'\
                 'IF(HESAP!B17<=0,"VERİ YOK",'\
                 'IF(AYARLAR!$B$17="Hayır",IF(HESAP!B49<=HESAP!B30,"PEŞİN ÖDE","TECİL YAP"),'\
                 'IF(HESAP!B50="TECİL","TECİL YAP",IF(HESAP!B50="PEŞİN","PEŞİN ÖDE","KREDİ KULLAN"))))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblBorc[Borç Aslı (TL)],"<0")>0,"Negatif borç aslı girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(COUNTIF(tblBorc[Gecikme Zammı (TL)],"<0")>0,"Negatif gecikme zammı girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(HESAP!B17<=0,"Borç kaydı bulunmuyor; ödeme yolu kararı üretilemiyor.",'\
                 'IF(AYARLAR!$B$17="Hayır",IF(HESAP!B49<=HESAP!B30,"Peşin ödeme toplam maliyeti tecilden düşük; borç peşin kapatılabilir.","Tecil toplam maliyeti peşin ödemenin altında; krediye erişim olmadığından tecil önerilir."),'\
                 'IF(HESAP!B50="TECİL","Tecil toplam maliyeti hem kredi hem peşin ödemeden düşük; tecil en avantajlı yoldur.",'\
                 'IF(HESAP!B50="PEŞİN","Peşin ödeme toplam maliyeti en düşük; terkin avantajıyla borç tek seferde kapatılabilir.",'\
                 '"Kredi toplam maliyeti en düşük; borç kredi ile kapatılıp vadeye yayılabilir."))))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=TecilAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki oran/eşiklerden türetilir; "
                     "tecil başvuru sonucu bağlayıcı değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Toplam Borç (TL)", "=HESAP!B17", TL),
        (3, "Tecil Toplam (TL)", "=HESAP!B30", TL),
        (4, "Kredi Toplam (TL)", "=HESAP!B39", TL),
        (5, "Peşin Ödeme (TL)", "=HESAP!B44", TL),
        (6, "En Düşük Seçenek", "=HESAP!B50", None),
        (7, "NBD Farkı (TL)", "=HESAP!B58", TL),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Kalite Skoru", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=KONTROLLER!B19", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "Borç Kalemi", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 8, "=HESAP!B20", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — borç türü bazlı
    h(ws, 10, 1, "Borç Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Asıl Toplam (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Zam Toplam (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Toplam Borç (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Aylık Zam Yükü (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, tur in enumerate(BORC_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!C{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!D{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=HESAP!E{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Borç kalemi bazlı toplam borç (kayıt sırasına göre) — tür etiket satırlarından ayrı
    h(ws, 20, 7, "Borç Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 20, 8, "Toplam Borç (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 21 + i, 7, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 21 + i, 8, f"=IFERROR(INDEX(tblBorc[Toplam Borç (TL)],AYARLAR!$B${34 + i}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Borç Türü Asıl ve Zam"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=17), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F8")
    g2 = PieChart()
    g2.title = "Toplam Borç Dağılımı"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=17), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Aylık Zam Yükü"
    g3.add_data(Reference(ws, min_col=5, min_row=10, max_row=17), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g3.height, g3.width = 9, 16
    ws.add_chart(g3, "F21")
    g4 = LineChart()
    g4.title = "Borç Kalemi Bazlı Toplam Borç Trendi"
    g4.add_data(Reference(ws, min_col=8, min_row=20, max_row=32), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=7, min_row=21, max_row=32))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U21")
    # Zaman serisi projeksiyonu
    h(ws, 34, 1, "Gelecek Dönem Toplam Borç Tahmini", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 34, 2, '=ROUND(IFERROR(FORECAST.LINEAR(TecilPanelSatir12,tblBorc[Toplam Borç (TL)],ROW(tblBorc[Toplam Borç (TL)])-ROW(tblBorc[[#Headers],[Toplam Borç (TL)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 35, 1, "Ortalama Toplam Borç", yazi="333333")
    h(ws, 35, 2, "=ROUND(IFERROR(AVERAGE(tblBorc[Toplam Borç (TL)]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 36, 1, "Toplam Borç Standart Sapması", yazi="333333")
    h(ws, 36, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblBorc[Toplam Borç (TL)],TecilPanelSatir1):INDEX(tblBorc[Toplam Borç (TL)],TecilPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 37, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 37, 2, "=B35-B36*TecilTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 38, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 38, 2, "=B35+B36*TecilTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=34, max_row=38), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=35, max_row=38))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F34")
    # Seçenek maliyet karşılaştırması
    h(ws, 40, 1, "Seçenek", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 40, 2, "Toplam Maliyet (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 40, 3, "NBD (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    secenekler = [("TECİL", "=HESAP!B30", "=HESAP!B55"),
                  ("KREDİ", "=HESAP!B39", "=HESAP!B56"),
                  ("PEŞİN", "=HESAP!B44", "=HESAP!B57")]
    for i, (ad, form, form2) in enumerate(secenekler, 41):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, form2, sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = BarChart()
    g6.type = "col"
    g6.title = "Seçenek Toplam Maliyet"
    g6.add_data(Reference(ws, min_col=2, min_row=40, max_row=43), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=41, max_row=43))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U34")
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Seçenek NBD Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=40, max_row=43), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=41, max_row=43))
    g7.height, g7.width = 8, 13
    ws.add_chart(g7, "F41")
    # Senaryo karşılaştırma verisi
    h(ws, 45, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 45, 6, "Tecil Maliyeti (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 46):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-39}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g8 = LineChart()
    g8.title = "Senaryo Tecil Maliyetleri"
    g8.add_data(Reference(ws, min_col=6, min_row=45, max_row=49), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=5, min_row=46, max_row=49))
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U41")
    baski_hazirla(ws, "A1:M50", URUN_AD)
    alt_bant(ws, 50, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 40, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Farklı tecil faizi senaryolarında toplam maliyet bant aralığı; ardından hangi "
                "değişkenin maliyeti en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Aylık Faiz", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Tecil Maliyeti (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "Durum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        ("İYİMSER", 20),
        ("BAZ", 21),
        ("KÖTÜMSER", 22),
        ("KRİTİK", 23),
    ]
    for i, (ad, ayar) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=AYARLAR!$B${ayar}/12", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "=IFERROR(HESAP!B25+HESAP!B26*(B" + str(i) + "/(1-(1+B" + str(i) + ")^-HESAP!B28))*HESAP!B28,0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(C{i}<=HESAP!B30,\"AVANTAJLI\",\"DEĞERLENDİR\")", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Senaryo Bant Genişliği (İyi − Kritik)", yazi="333333", kalin=True)
    h(ws, 11, 3, "=C7-C10", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Tornado: Toplam Maliyete Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Tecil Faizi +1 puan Etkisi", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ABS((HESAP!B25+HESAP!B26*((AYARLAR!$B$13+TecilTornadoFaiz)/12)/(1-(1+(AYARLAR!$B$13+TecilTornadoFaiz)/12)^-HESAP!B28)*HESAP!B28)-HESAP!B30),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Kredi Faizi +1 puan Etkisi", yazi="333333")
    h(ws, 15, 2, "=IFERROR(ABS((HESAP!B34+HESAP!B35*((AYARLAR!$B$14+KrediTornadoFaiz)/12)/(1-(1+(AYARLAR!$B$14+KrediTornadoFaiz)/12)^-HESAP!B37)*HESAP!B37)-HESAP!B39),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, "=IF(B14>=B15,\"Tecil Faizi\",\"Kredi Faizi\")", yazi=DIS_REF_YEŞIL, kalin=True)
    g = BarChart()
    g.type = "bar"
    g.title = "Tornado: Maliyete Etki"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 8, 13
    ws.add_chart(g, "E17")
    alt_bant(ws, 19, "Senaryo faizleri ve tornado çarpanları AYARLAR'dan türetilir.")
    genislik(ws, {"A": 42, "B": 30, "C": 22, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "VERGİ/SGK BORCU — ÖDEME YOLU YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Toplam Borç (TL)", "=HESAP!B17", TL),
        ("Toplam Gecikme Zammı (TL)", "=HESAP!B16", TL),
        ("Tecil Toplam Maliyet (TL)", "=HESAP!B30", TL),
        ("Kredi Toplam Maliyet (TL)", "=HESAP!B39", TL),
        ("Peşin Ödeme (TL)", "=HESAP!B44", TL),
        ("En Düşük Maliyet (TL)", "=HESAP!B51", TL),
        ("NBD Farkı (Tecil − Kredi) (TL)", "=HESAP!B58", TL),
        ("Ortalama Geciken Ay", "=HESAP!B19", "0.0"),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "Tecil seçeneği 6183 sayılı Kanun 48. madde kapsamında vadeye yaymayı modeller; tecil faizi AYARLAR'dan güncellenir.",
        "Kredi seçeneği borcu kapatacak banka kredisini modeller; faiz ve komisyon oranı kullanıcı tarafından girilir.",
        "Peşin ödeme seçeneği yapılandırma kapsamında gecikme zammının terkin edilen kısmını dikkate alır.",
        "NBD karşılaştırması iskonto oranıyla aylık ödemeleri bugünkü değere indirger; iskonto oranı AYARLAR'dadır.",
        "Bu dosya karar destek aracıdır; tecil başvuru sonucu ve kredi onayı bağlayıcı değildir.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=TecilAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
        ws.merge_cells(start_row=25 + i, start_column=1, end_row=25 + i, end_column=12)
    h(ws, 29, 1, "Bu rapor karar destek amaçlıdır; mali tavsiye niteliği taşımaz.",
      yazi=GRİ, boyut=9, italik=True, kaydir=True)
    baski_hazirla(ws, "A1:L29", URUN_AD)
    genislik(ws, {"A": 45, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek Veri (Demo)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Gerçek girişlere başlamadan önce demo veri ile ürünü deneyin. "
                "Bu satırlar örnektir; gerçek verilerle değiştirin.", yazi=GRİ, kaydir=True)
    sutunlar = ["Borç Kodu", "Borç Türü", "Borç Aslı (TL)", "Gecikme Zammı (TL)",
                "Vade Tarihi", "Aylık Zam Oranı", "Toplam Borç (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Toplam Borç (TL)": "=IF(tblOrnekBorc[[#This Row],[Borç Aslı (TL)]]=\"\",0,"\
                            "tblOrnekBorc[[#This Row],[Borç Aslı (TL)]]+tblOrnekBorc[[#This Row],[Gecikme Zammı (TL)]])",
        "Açıklama": "=IF(tblOrnekBorc[[#This Row],[Borç Türü]]=\"\",\"\",CONCATENATE(\"Örnek borç: \",tblOrnekBorc[[#This Row],[Borç Türü]]))",
    }
    tablo_ekle(ws, "tblOrnekBorc", "A5:H1004", sutunlar, formuller)
    ornek = [
        ("V-1001", "KDV", 150000, 42000, date(2025, 3, 26), 0.04),
        ("V-1002", "Muhtasar Stopaj", 80000, 19000, date(2025, 6, 26), 0.04),
        ("V-1003", "Kurumlar Vergisi", 240000, 61000, date(2025, 5, 31), 0.04),
        ("S-2001", "SGK Primi", 120000, 31000, date(2025, 8, 31), 0.03),
        ("S-2002", "SGK Cezası", 40000, 14000, date(2025, 10, 31), 0.04),
        ("V-1004", "Gelir Vergisi", 60000, 9000, date(2025, 11, 30), 0.04),
    ]
    giris_kolonlari = [1, 2, 3, 4, 5, 6]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir):
            h(ws, i, giris_kolonlari[k], deger)
        ws.cell(row=i, column=3).number_format = TL
        ws.cell(row=i, column=4).number_format = TL
        ws.cell(row=i, column=5).number_format = TARİH
        ws.cell(row=i, column=6).number_format = YÜZDE
        ws.cell(row=i, column=7).number_format = TL
    genislik(ws, {"A": 11, "B": 20, "C": 18, "D": 18, "E": 13, "F": 14, "G": 18, "H": 24})
    alt_bant(ws, 1006, "Demo veriler rapor sonuçlarını örnekler; satın alma sonrası silinerek "
                       "gerçek veriler girilir.")


def degisiklik_kaydi(ws):
    sayfa_hazirla(ws, "DEGISIKLIK_KAYDI", "999999", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Değişiklik Kaydı", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 4, ["Tarih", "Sürüm", "Değişiklik", "Neden", "Yapan"])
    kayitlar = [
        ("10.08.2026", SURUM, "İlk üretim", "Ürünün yayına hazırlanması", "ExcelArşiv"),
    ]
    for i, kayit in enumerate(kayitlar, 5):
        for k, deger in enumerate(kayit, 1):
            h(ws, i, k, deger)
    genislik(ws, {"A": 14, "B": 10, "C": 40, "D": 30, "E": 14})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "808080", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Bağlı Listeler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Bu sayfadaki listeler açılır menüleri besler; yeni seçenek eklemek için "
                "sarı hücrelere yazın.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Borç Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(BORC_TURLERI, 7):
        h(ws, i, 1, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Borç türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni tür için satır ekleyin")
    h(ws, 7, 2, "Kredi Onayı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, sec in enumerate(["Evet", "Hayır"], 8):
        h(ws, i, 2, sec, zemin=GIRIS_SARI)
    genislik(ws, {"A": 18, "B": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Tecil faizi, kredi faizi, iskonto oranı ve terkin "
                "oranını güncelleyin; değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"])
    TARIH_D = date(2026, 8, 10)
    YURURLUK = date(2025, 1, 1)
    DOGRULAMA = date(2026, 8, 10)
    satirlar = [
        (7, "FirmaUnvan", "Örnek Firma A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "RaporTarihi", TARIH_D, "Tarih", None, None, "Geciken ay ve NBD hesabının dayanak tarihi.", "Kullanıcı"),
        (9, "RaporHazirlayan", "Mali İşler Direktörü", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "DosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "TecilPesinatOran", 0.10, "Oran", 0, 1, "Tecile başlamadan ödenen peşin oranı.", "Varsayım"),
        (12, "TecilTaksitSayisi", 12, "Adet", 1, 60, "Tecilin vadeye yayılacağı taksit sayısı.", "Varsayım"),
        (13, "TecilFaizYillik", 0.18, "Oran", 0, 1, "Yıllık tecil faizi (6183 kapsamında değişken).", "Varsayım"),
        (14, "KrediFaizYillik", 0.35, "Oran", 0, 1, "Borcu kapatacak banka kredisinin yıllık faizi.", "Varsayım"),
        (15, "KrediKomisyonOran", 0.02, "Oran", 0, 1, "Kredi komisyon oranı.", "Varsayım"),
        (16, "KrediTaksitSayisi", 12, "Adet", 1, 60, "Kredi taksit sayısı.", "Varsayım"),
        (17, "KrediOnayiVar", "Evet", "Metin", None, None, "Kredi onayı var mı? Hayır ise kredi seçeneği devre dışı kalır.", "Varsayım"),
        (18, "PesinTerkinOran", 0.50, "Oran", 0, 1, "Peşin ödemede gecikme zammından terkin edilen oran (yapılandırma).", "Varsayım"),
        (19, "IskontoOrani", 0.30, "Oran", 0, 1, "NBD için yıllık iskonto oranı.", "Varsayım"),
        (20, "SenTecilIyi", 0.12, "Oran", 0, 1, "İyimser senaryo tecil faizi.", "Varsayım"),
        (21, "SenTecilBaz", 0.18, "Oran", 0, 1, "Baz senaryo tecil faizi.", "Varsayım"),
        (22, "SenTecilKotu", 0.25, "Oran", 0, 1, "Kötümser senaryo tecil faizi.", "Varsayım"),
        (23, "SenTecilKritik", 0.35, "Oran", 0, 1, "Kritik senaryo tecil faizi.", "Varsayım"),
        (24, "TecilTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (25, "TecilKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (26, "TecilKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (27, "TecilAylikNakitAkisi", 100000, "TL", 0, 100000000000, "Aylık nakit akışı; tecil taksiti bu değerle kıyaslanır.", "Varsayım"),
        (28, "TecilYuksekBorcEsik", 500000, "TL", 0, 100000000000, "Yüksek borç eşiği.", "Varsayım"),
        (29, "TecilP90Oran", 0.90, "Oran", 0, 1, "Toplam borç dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (30, "TecilTemelKolonSayisi", 5, "Adet", 1, 20, "Veri kalitesi skorunda temel kabul edilen zorunlu giriş kolonu sayısı.", "Varsayım"),
        (31, "TecilTornadoFaiz", 0.01, "Oran", 0, 0.1, "Tornado analizinde tecil faizine uygulanan +1 puan çarpanı.", "Varsayım"),
        (32, "KrediTornadoFaiz", 0.01, "Oran", 0, 0.1, "Tornado analizinde kredi faizine uygulanan +1 puan çarpanı.", "Varsayım"),
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
        ws.row_dimensions[satir].height = 26
    dogrulama(ws, "list", "ListeEvetHayir", "B17",
              baslik="Kredi Onayı", mesaj="Kredi onayı durumunu seçin; Hayır ise karar kapısı kredi seçeneğini dışarıda bırakır.",
              hata_baslik="Geçersiz seçim", hata_mesaj="Listeden Evet veya Hayır seçmelisiniz.")
    for i in range(1, 13):
        satir = 34 + i
        h(ws, satir, 1, f"TecilPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki borç kalemi bazlı toplam borç grafiğinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki grafiği besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    for i in range(3):
        satir = 48 + i
        h(ws, satir, 1, f"TecilAksiyon{i+1}_ACIKLAMA", kalin=True, zemin=ACIK_GRI, boyut=9)
        h(ws, satir, 2, "Karara göre önerilen aksiyon.", zemin=GIRIS_SARI, yazi=GIRIS_YAZI, kaydir=True)
        h(ws, satir, 3, "Metin", boyut=9, yazi=GRİ)
        h(ws, satir, 6, "KARAR sayfasında gösterilen aksiyon metni.", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Varsayım", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        ws.row_dimensions[satir].height = 30
    genislik(ws, {"A": 32, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 20, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "BORC_LISTESI sayfasında sarı hücrelere borçlarınızı girin. Borç türünü listeden "
         "seçin; borç aslı, gecikme zammı, vade tarihi ve aylık zam oranını girin. "
         "Geciken ay, toplam borç ve aylık zam yükü otomatik hesaplanır."),
        ("Seçenek hesapları",
         "HESAP sayfası tecil, kredi ve peşin ödeme seçeneklerinin toplam maliyetini ve "
         "bugünkü değerini (NBD) hesaplar. Peşinat oranı, taksit sayısı ve faizler AYARLAR'dan gelir."),
        ("Karar ve aksiyon",
         "KARAR sayfası en düşük maliyetli ödeme yolunu gerekçesiyle önerir. Kredi onayı "
         "yoksa kredi seçeneği karar dışı bırakılır."),
        ("Oranlar ve senaryolar",
         "AYARLAR sayfasından tecil faizi, kredi faizi, iskonto oranı, terkin oranı ve senaryo "
         "faizlerini güncelleyin; PANO ve SENARYO_DUYARLILIK sayfaları anında güncellenir."),
        ("Güven ve karar kalitesi",
         "Veri kalite skoru girişlerin tamlığını ölçer. Borç girişi yokken karar VERİ YOK; "
         "negatif girişte DURDUR üretir. Verileriniz cihazınızdan çıkmaz; dosya tamamen çevrimdışıdır."),
        ("Uyumluluk",
         "Excel 2016-365 (Windows ve Mac), LibreOffice ve Google Sheets ile uyumludur. "
         "Makro içermez; formüller ayrıştırıcı ile denetlenmiştir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        ws.row_dimensions[i].height = 40
    genislik(ws, {"A": 34, "B": 110})


def kosullu_bicimlendirme(wb):
    kirmizi = PatternFill("solid", start_color="FDE9E9", end_color="FDE9E9")
    yesil = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
    sari = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    koyu = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
    kirmizi_font = Font(name=FONT, color="B3261E", bold=True, size=10)
    yesil_font = Font(name=FONT, color="1F7A4D", bold=True, size=10)

    ws = wb["KARAR"]
    for hucre in ("B11",):
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="equal", formula=['"TECİL YAP"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="equal", formula=['"KREDİ KULLAN"'], fill=sari))
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="equal", formula=['"PEŞİN ÖDE"'], fill=yesil, font=yesil_font))

    ws = wb["PANO"]
    for r in (4, 6):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
    for r in range(11, 18):
        for c in ("B", "C", "D", "E"):
            ws.conditional_formatting.add(f"{c}{r}",
                CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    for c in ("B", "C", "D", "E"):
        ws.conditional_formatting.add(f"{c}11:{c}17",
            CellIsRule(operator="greaterThan", formula=["500000"], fill=sari))
    ws.conditional_formatting.add("B11:B22",
        CellIsRule(operator="greaterThan", formula=["100000"], fill=koyu))

    ws = wb["HESAP"]
    for hucre in ("B29", "B38", "B44", "B50", "B51", "B58"):
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B17",
        CellIsRule(operator="greaterThan", formula=["TecilYuksekBorcEsik"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B30",
        CellIsRule(operator="lessThanOrEqual", formula=["B39"], fill=yesil))
    ws.conditional_formatting.add("B39",
        CellIsRule(operator="lessThan", formula=["B30"], fill=sari))
    for r in range(6, 13):
        ws.conditional_formatting.add(f"D{r}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

    ws = wb["BORC_LISTESI"]
    ws.conditional_formatting.add("C6:C1004",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("D6:D1004",
        CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("G6:G1004",
        CellIsRule(operator="greaterThan", formula=["36"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("H6:H1004",
        CellIsRule(operator="greaterThan", formula=["200000"], fill=sari))
    ws.conditional_formatting.add("F6:F1004",
        CellIsRule(operator="greaterThan", formula=["1"], fill=kirmizi, font=kirmizi_font))

    ws = wb["KONTROLLER"]
    for r in range(6, 14):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="notEqual", formula=['"TEMİZ"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B19",
        CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yesil_font))
    ws.conditional_formatting.add("B19",
        CellIsRule(operator="between", formula=["70", "89"], fill=sari))
    ws.conditional_formatting.add("B19",
        CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kirmizi_font))

    ws = wb["SENARYO_DUYARLILIK"]
    for r in range(7, 11):
        ws.conditional_formatting.add(f"C{r}",
            CellIsRule(operator="lessThanOrEqual", formula=["$C$7"], fill=yesil))
        ws.conditional_formatting.add(f"C{r}",
            CellIsRule(operator="greaterThan", formula=["$C$7"], fill=sari))
    ws.conditional_formatting.add("C11",
        CellIsRule(operator="greaterThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

    ws = wb["RAPOR"]
    ws.conditional_formatting.add("B4",
        CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B4",
        CellIsRule(operator="notEqual", formula=['"DURDUR"'], fill=yesil, font=yesil_font))


def main(cikti_yolu=None):
    wb = Workbook()
    ilk = wb.active
    sayfalar = [
        (kapak, "KAPAK"), (hizli_baslangic, "HIZLI_BASLANGIC"),
        (borc_listesi, "BORC_LISTESI"), (hesap, "HESAP"),
        (analitik_motor, "ANALITIK_MOTOR"), (kontroller, "KONTROLLER"),
        (karar_motoru, "KARAR"), (pano, "PANO"),
        (senaryo_duyarlilik, "SENARYO_DUYARLILIK"), (rapor, "RAPOR"),
        (ornek_veri, "ORNEK_VERI"), (degisiklik_kaydi, "DEGISIKLIK_KAYDI"),
        (listeler, "LISTELER"), (ayarlar, "AYARLAR"), (kilavuz, "KILAVUZ"),
    ]
    for fn, ad in sayfalar:
        ws = ilk if ad == "KAPAK" else wb.create_sheet()
        fn(ws)

    # Ad tanımları — motor çıktıları
    ad_ekle(wb, "TecilToplamBorc", "HESAP!$B$17")
    ad_ekle(wb, "TecilToplamMaliyet", "HESAP!$B$30")
    ad_ekle(wb, "KrediToplamMaliyet", "HESAP!$B$39")
    ad_ekle(wb, "PesinOdeme", "HESAP!$B$44")
    ad_ekle(wb, "EnDusukSecenek", "HESAP!$B$50")
    ad_ekle(wb, "NbdFarki", "HESAP!$B$58")
    ad_ekle(wb, "TecilAylikTaksit", "HESAP!$B$29")
    ad_ekle(wb, "KrediAylikTaksit", "HESAP!$B$38")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "TecilPesinatOran": 11, "TecilTaksitSayisi": 12, "TecilFaizYillik": 13,
        "KrediFaizYillik": 14, "KrediKomisyonOran": 15, "KrediTaksitSayisi": 16,
        "KrediOnayiVar": 17, "PesinTerkinOran": 18, "IskontoOrani": 19,
        "SenTecilIyi": 20, "SenTecilBaz": 21, "SenTecilKotu": 22, "SenTecilKritik": 23,
        "TecilTahminCarpan": 24, "TecilKaliteIyi": 25, "TecilKaliteOrta": 26,
        "TecilAylikNakitAkisi": 27, "TecilYuksekBorcEsik": 28, "TecilP90Oran": 29,
        "TecilTemelKolonSayisi": 30, "TecilTornadoFaiz": 31, "KrediTornadoFaiz": 32,
    }
    for i in range(1, 13):
        ayar_satirlari[f"TecilPanelSatir{i}"] = 34 + i
    for i in range(3):
        ayar_satirlari[f"TecilAksiyon{i+1}_ACIKLAMA"] = 48 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "TecilToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "TecilDoluGiris", "KONTROLLER!$B$18")
    ad_ekle(wb, "TecilRaporTarihi", "AYARLAR!$B$8")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1ZamOrani": 6, "modulT1ZamOraniYorum": 6,
        "modulT2OrtalamaBorc": 7, "modulT2OrtalamaBorcYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2EnYuksekBorc": 9, "modulO2EnYuksekBorcYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1BorcTahmin": 13, "modulI1BorcTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
        "modulI3NbdFarki": 15, "modulI3NbdFarkiYorum": 15,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeBorcTurleri", "LISTELER!$A$7:$A$13")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$B$8:$B$9")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "VergiSgkBorcunuTecilEtmeliMiyimKrediMiTecilMi.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
