"""
Amortisman & Sabit Kıymet Satış Zamanlama Stratejisti — üretim betiği.
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

URUN_AD = "Amortisman & Sabit Kıymet Satış Zamanlama Stratejisti"
SURUM = "1.0.0"
RENK = "8064A2"

KIYMET_TURLERI = ["Bina", "Makine", "Taşıt", "Demirbaş", "Yazılım", "Diğer"]
YONTEMLER = ["Normal", "Azalan"]
CEYREKLER = ["1. ÇEYREK", "2. ÇEYREK", "3. ÇEYREK", "4. ÇEYREK"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Sabit kıymetlerinizi amortisman yöntemi ve satış bedeliyle girin; "
                "yıl içinde dört çeyreğe göre satışın birikmiş amortisman, net defter "
                "değeri, kâr/zarar, vergi sonrası net nakit ve bugünkü değerini "
                "karşılaştırın; satış için en uygun döneme karar verin.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Kıymet bazında normal ve azalan bakiyeli amortisman, birikmiş amortisman ve net defter değeri hesabı",
        "Yıl içinde dört çeyrek satış senaryosunda ayrılan kıst amortisman, satış kârı/zararı, kurumlar vergisi etkisi ve net nakit hesabı",
        "Çeyrek nakit akışlarını bugünkü değerine indirgeyerek satış için en uygun dönemi önerir",
        "Satış bedeli ve amortisman oranına göre senaryo bant aralığı ve tornado analizi",
        "Kıymet türü yoğunlaşması, anomali ve veri kalite skoru ile karar kapısı ve aksiyon önerisi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Sabit kıymet satışının vergi etkisini hesaplamak isteyen işletmeler",
        "Amortisman planlaması ve satış zamanlaması yapan finans yöneticileri",
        "Müşterisine kıymet bazlı satış/amortisman stratejisi öneren mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. KIYMET_LISTESI sayfasına kıymetlerinizi girin. "
                 "3. PANO ve RAPOR sayfalarından satış için en uygun dönemi izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Kıymetlerinizi girin",
         "KIYMET_LISTESI sayfasındaki sarı hücrelere kıymet kodu, kıymet türü, edinim "
         "tarihi, edinim bedeli, amortisman yöntemi ve oranı ile satış bedeli tahminini girin."),
        ("Adım 2 — Çeyrek senaryolarını izleyin",
         "HESAP sayfası dört çeyrek satış senaryosunun kıst amortisman, kâr/zarar, "
         "vergi etkisi ve net nakit değerlerini hesaplar."),
        ("Adım 3 — Kararı değerlendirin",
         "KARAR sayfası net nakit ve NBD'ye göre en uygun satış dönemini önerir."),
        ("Adım 4 — Oranları ayarlayın",
         "Kurumlar vergisi oranı, iskonto oranı, azalan bakiye çarpanı ve kıst oranlar AYARLAR sayfasındadır."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Azalan bakiyeli kıymette oranı 0,50'nin üzerinde girmek (denge bozulur)",
        "Edinim tarihini rapor tarihinden sonra girmek (geçen yıl sıfır görünür)",
        "Satış bedelini net defter değerine eşitlemek (kâr/zarar görünmez olur)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; amortisman ve vergi hesabı mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def kIymet_listesi(ws):
    sayfa_hazirla(ws, "KIYMET_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sabit Kıymet Listesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Her satır bir sabit kıymettir; "
                "geçen yıl, birikmiş amortisman, net defter değeri, yıllık amortisman "
                "ve satış kârı/zararı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kıymet Kodu", "Kıymet Türü", "Edinim Tarihi", "Edinim Bedeli (TL)",
                "Amortisman Yöntemi", "Amortisman Oranı", "Satış Bedeli (TL)",
                "Geçen Yıl", "Birikmiş Amortisman (TL)", "Net Defter Değeri (TL)",
                "Yıllık Amortisman (TL)", "Satış Kârı/Zararı (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Geçen Yıl": '=IF(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                     'MAX(0,ROUNDDOWN((AmortismanRaporTarihi-tblKiyamet[[#This Row],[Edinim Tarihi]])/365,0)))',
        "Birikmiş Amortisman (TL)": '=IF(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                                    'IFERROR(IF(tblKiyamet[[#This Row],[Amortisman Yöntemi]]="Azalan",'
                                    'MAX(0,tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]*(1-MIN(2*AmortismanAzalanCarpan*'
                                    'tblKiyamet[[#This Row],[Amortisman Oranı]],1)^tblKiyamet[[#This Row],[Geçen Yıl]])),'
                                    'MIN(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]],tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]*'
                                    'tblKiyamet[[#This Row],[Amortisman Oranı]]*tblKiyamet[[#This Row],[Geçen Yıl]])),0))',
        "Net Defter Değeri (TL)": '=IF(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                                  'MAX(0,tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]-tblKiyamet[[#This Row],[Birikmiş Amortisman (TL)]]))',
        "Yıllık Amortisman (TL)": '=IF(tblKiyamet[[#This Row],[Geçen Yıl]]=0,0,'
                                  'IFERROR(IF(tblKiyamet[[#This Row],[Amortisman Yöntemi]]="Azalan",'
                                  'tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]*(1-MIN(2*AmortismanAzalanCarpan*'
                                  'tblKiyamet[[#This Row],[Amortisman Oranı]],1))^MAX(0,tblKiyamet[[#This Row],[Geçen Yıl]]-1)*'
                                  'MIN(2*AmortismanAzalanCarpan*tblKiyamet[[#This Row],[Amortisman Oranı]],1),'
                                  'MIN(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]*tblKiyamet[[#This Row],[Amortisman Oranı]],'
                                  'tblKiyamet[[#This Row],[Edinim Bedeli (TL)]])),0))',
        "Satış Kârı/Zararı (TL)": '=IF(tblKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                                  'tblKiyamet[[#This Row],[Satış Bedeli (TL)]]-tblKiyamet[[#This Row],[Net Defter Değeri (TL)]])',
        "Açıklama": '=IF(tblKiyamet[[#This Row],[Kıymet Türü]]="","",'
                    'CONCATENATE("Kıymet: ",tblKiyamet[[#This Row],[Kıymet Türü]]," | NDD: ",TEXT(tblKiyamet[[#This Row],[Net Defter Değeri (TL)]],"#,##0")))',
    }
    tablo_ekle(ws, "tblKiyamet", "A5:M1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeKiyametTurleri", "B6:B1004",
              baslik="Kıymet Türü", mesaj="Listeden bir kıymet türü seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "date", "01.01.2000", "C6:C1004",
              baslik="Edinim Tarihi", mesaj="Kıymetin edinim tarihini girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2000 ile 31.12.2100 arasında olmalı.",
              isaret="between", f2="31.12.2100")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Edinim Bedeli", mesaj="Kıymetin edinim bedelini TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "list", "ListeAmortismanYontemleri", "E6:E1004",
              baslik="Amortisman Yöntemi", mesaj="Normal veya Azalan seçin.",
              hata_baslik="Geçersiz yöntem", hata_mesaj="Normal veya Azalan seçmelisiniz.")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Amortisman Oranı", mesaj="Yıllık amortisman oranını 0-0,50 arasında girin (örn. %20 için 0,20).",
              hata_baslik="Geçersiz oran", hata_mesaj="Oran 0 ile 0,50 arasında olmalı.",
              isaret="between", f2="0.5")
    dogrulama(ws, "decimal", "0", "G6:G1004",
              baslik="Satış Bedeli", mesaj="Kıymetin tahmini satış bedelini TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=6).number_format = YÜZDE
        ws.cell(row=satir, column=7).number_format = TL
    ornek = [
        ("K-1001", "Bina", date(2019, 5, 15), 1500000, "Normal", 0.02, 1900000),
        ("K-1002", "Makine", date(2021, 3, 10), 600000, "Azalan", 0.20, 320000),
        ("K-1003", "Makine", date(2020, 9, 1), 400000, "Normal", 0.20, 210000),
        ("K-1004", "Taşıt", date(2022, 1, 20), 350000, "Normal", 0.20, 190000),
        ("K-1005", "Demirbaş", date(2023, 4, 5), 180000, "Normal", 0.20, 95000),
        ("K-1006", "Yazılım", date(2024, 6, 15), 120000, "Azalan", 0.33, 60000),
        ("K-1007", "Bina", date(2016, 11, 1), 900000, "Normal", 0.02, 1250000),
    ]
    giris_kolonlari = [1, 2, 3, 4, 5, 6, 7]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir):
            ws.cell(row=i, column=giris_kolonlari[k]).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7])
    sabitle(ws, "A6")
    genislik(ws, {"A": 13, "B": 18, "C": 13, "D": 18, "E": 18, "F": 15, "G": 18,
                  "H": 11, "I": 22, "J": 22, "K": 20, "L": 20, "M": 30})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kıymet Türü Bazlı Özet ve Satış Zamanlama Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet kıymet listesinden otomatik beslenir; çeyrek satış senaryoları "
                "AYARLAR'daki oran ve kıst değerlerle canlı üretilir.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kıymet Türü", "Adet", "Edinim Bedeli (TL)", "Birikmiş Amortisman (TL)",
                "Net Defter Değeri (TL)", "Yıllık Amortisman (TL)", "Satış Bedeli (TL)"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(KIYMET_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=IFERROR(COUNTIFS(tblKiyamet[Kıymet Türü],$A{r},tblKiyamet[Edinim Bedeli (TL)],"<>0"),0)',
          sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblKiyamet[Edinim Bedeli (TL)],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblKiyamet[Birikmiş Amortisman (TL)],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblKiyamet[Net Defter Değeri (TL)],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=SUMIFS(tblKiyamet[Yıllık Amortisman (TL)],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f'=SUMIFS(tblKiyamet[Satış Bedeli (TL)],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "TOPLAM BLOK", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Toplam Edinim Bedeli (TL)", yazi="333333")
    h(ws, 14, 2, "=SUM(C6:C11)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 15, 1, "Toplam Birikmiş Amortisman (TL)", yazi="333333")
    h(ws, 15, 2, "=SUM(D6:D11)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "TOPLAM NET DEFTER DEĞERİ (TL)", yazi="333333", kalin=True)
    h(ws, 16, 2, "=SUM(E6:E11)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 17, 1, "Toplam Yıllık Amortisman (TL)", yazi="333333")
    h(ws, 17, 2, "=SUM(F6:F11)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Toplam Satış Bedeli (TL)", yazi="333333")
    h(ws, 18, 2, "=SUM(G6:G11)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kıymet Sayısı", yazi="333333")
    h(ws, 19, 2, "=COUNTA(tblKiyamet[Kıymet Kodu])", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 20, 1, "Kıymet Türü Sayısı (farklı)", yazi="333333")
    h(ws, 20, 2, '=IFERROR(SUMPRODUCT((tblKiyamet[Kıymet Türü]<>"")/COUNTIF(tblKiyamet[Kıymet Türü],tblKiyamet[Kıymet Türü]&"")),0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    # Çeyrek satış senaryoları
    bloklar = [
        ("1. ÇEYREK SATIŞ SENARYOSU (mart sonu)", 22, 14),
        ("2. ÇEYREK SATIŞ SENARYOSU (haziran sonu)", 30, 15),
        ("3. ÇEYREK SATIŞ SENARYOSU (eylül sonu)", 38, 16),
        ("4. ÇEYREK SATIŞ SENARYOSU (aralık sonu)", 46, 17),
    ]
    for baslik, b0, kist in bloklar:
        h(ws, b0, 1, baslik, kalin=True, boyut=11, yazi=KOYU_LACIVERT)
        ws.merge_cells(f"A{b0}:G{b0}")
        h(ws, b0 + 1, 1, "Satış Yılı Kıst Amortismanı (TL)", yazi="333333")
        h(ws, b0 + 1, 2, f"=B17*AYARLAR!$B${kist}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, b0 + 2, 1, "Toplam Birikmiş Amortisman (TL)", yazi="333333")
        h(ws, b0 + 2, 2, f"=B15+B{b0 + 1}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, b0 + 3, 1, "Net Defter Değeri (TL)", yazi="333333")
        h(ws, b0 + 3, 2, f"=B14-B{b0 + 2}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, b0 + 4, 1, "Satış Kârı/Zararı (TL)", yazi="333333")
        h(ws, b0 + 4, 2, f"=B18-B{b0 + 3}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, b0 + 5, 1, "Kurumlar Vergisi Etkisi (TL)", yazi="333333")
        h(ws, b0 + 5, 2, f"=B{b0 + 4}*AmortismanKurumlarVergiOran", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, b0 + 6, 1, "NET NAKİT ETKİ (TL)", yazi="333333", kalin=True)
        h(ws, b0 + 6, 2, f"=B18-B{b0 + 5}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    # Karşılaştırma
    h(ws, 55, 1, "ÇEYREK KARŞILAŞTIRMASI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A55:G55")
    h(ws, 56, 1, "Net Nakit — 1Ç (TL)", yazi="333333")
    h(ws, 56, 2, "=B28", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 57, 1, "Net Nakit — 2Ç (TL)", yazi="333333")
    h(ws, 57, 2, "=B36", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 58, 1, "Net Nakit — 3Ç (TL)", yazi="333333")
    h(ws, 58, 2, "=B44", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 59, 1, "Net Nakit — 4Ç (TL)", yazi="333333")
    h(ws, 59, 2, "=B52", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 60, 1, "EN YÜKSEK NET NAKİT (TL)", yazi="333333", kalin=True)
    h(ws, 60, 2, "=MAX(B56:B59)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    # NBD zamanlaması
    h(ws, 62, 1, "NBD — 1Ç (TL)", yazi="333333")
    h(ws, 62, 2, "=B28/(1+AmortismanIskontoOrani)^AYARLAR!$B$14", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 63, 1, "NBD — 2Ç (TL)", yazi="333333")
    h(ws, 63, 2, "=B36/(1+AmortismanIskontoOrani)^AYARLAR!$B$15", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 64, 1, "NBD — 3Ç (TL)", yazi="333333")
    h(ws, 64, 2, "=B44/(1+AmortismanIskontoOrani)^AYARLAR!$B$16", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 65, 1, "NBD — 4Ç (TL)", yazi="333333")
    h(ws, 65, 2, "=B52/(1+AmortismanIskontoOrani)^AYARLAR!$B$17", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 66, 1, "EN İYİ NBD ÇEYREĞİ", yazi="333333", kalin=True)
    h(ws, 66, 2, '=IF(B62>=B63,IF(B62>=B64,IF(B62>=B65,"1. ÇEYREK","4. ÇEYREK"),IF(B64>=B65,"3. ÇEYREK","4. ÇEYREK")),IF(B63>=B64,IF(B63>=B65,"2. ÇEYREK","4. ÇEYREK"),IF(B64>=B65,"3. ÇEYREK","4. ÇEYREK")))',
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 67, 1, "EN İYİ NBD DEĞERİ (TL)", yazi="333333", kalin=True)
    h(ws, 67, 2, "=MAX(B62:B65)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 68, 1, "NBD Farkı (1Ç − 4Ç) (TL)", yazi="333333")
    h(ws, 68, 2, "=B62-B65", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 70, "Hesap satırları yalnızca kıymet listesinden ve AYARLAR'dan beslenir; korumalıdır.")
    genislik(ws, {"A": 42, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20})
def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Ortalama Amortisman Oranı",
         "=IFERROR(AVERAGE(tblKiyamet[Amortisman Oranı]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama amortisman oranı ",TEXT(C6,"0.0%")," — kıymet portföyünün ortalama amortisman hızını gösterir")'),
        ("T2", "Ortalama Net Defter Değeri",
         "=IFERROR(AVERAGE(tblKiyamet[Net Defter Değeri (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama net defter değeri ",TEXT(C7,"₺ #,##0")," TL — portföyün ortalama kıymet büyüklüğünü gösterir")'),
        ("O1", "NDD Sapması (MAD)",
         "=IFERROR(AVEDEV(tblKiyamet[Net Defter Değeri (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Net defter değeri ortalama mutlak sapması ",TEXT(C8,"₺ #,##0")," TL — sapma yüksekse kıymetler heterojendir")'),
        ("O2", "En Yüksek Satış Kârı/Zararı",
         "=IFERROR(MAX(tblKiyamet[Satış Kârı/Zararı (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek satış kârı ",TEXT(C9,"₺ #,##0")," TL — bu kıymet satış zamanlamasında önceliklidir")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C11",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kritik senaryo net nakit farkı ",TEXT(C10,"₺ #,##0")," TL — satış bedeli belirsizliğinin etkisini gösterir")'),
        ("O6", "Edinim Yoğunlaşması (HHI)",
         "=IFERROR(IF(SUM(HESAP!$C$6:$C$11)=0,0,SUMPRODUCT((HESAP!$C$6:$C$11/SUM(HESAP!$C$6:$C$11))^2)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Edinim yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça yatırım tek türde toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=AmortismanKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Amortisman Tahmini",
         "=PANO!B35",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem toplam amortisman tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "NDD Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblKiyamet[Net Defter Değeri (TL)],AmortismanP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kıymetlerin %90 ı ",TEXT(C14,"₺ #,##0")," TL altında net defter değerine sahiptir — üst sınır planlamada kullanılabilir")'),
        ("I3", "NBD Farkı (1Ç − 4Ç)",
         "=HESAP!B68",
         '=_xlfn.TEXTJOIN("; ",TRUE,"NBD farkı ",TEXT(C15,"₺ #,##0")," TL — fark pozitifse erken satış bugünkü değerde avantajlıdır")'),
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
        ("Boş edinim bedeli satırı var mı?", "=IF(COUNTBLANK(tblKiyamet[Edinim Bedeli (TL)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, "B3261E"),
        ("Negatif edinim bedeli var mı?", "=IF(COUNTIF(tblKiyamet[Edinim Bedeli (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Negatif satış bedeli var mı?", "=IF(COUNTIF(tblKiyamet[Satış Bedeli (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Amortisman oranı 0,50 üzerinde mi?", "=IF(COUNTIF(tblKiyamet[Amortisman Oranı],\">0.5\")=0,\"TEMİZ\",\"ORAN >0,5 VAR\")", NORMAL, "B3261E"),
        ("Edinimi 40 yıldan eski kıymet var mı?", "=IF(COUNTIF(tblKiyamet[Geçen Yıl],\">40\")=0,\"TEMİZ\",\"ESKİ VAR\")", NORMAL, "B3261E"),
        ("NDD eşiği aşılıyor mu?", "=IF(HESAP!B16>AmortismanYuksekNddEsik,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("En iyi çeyrek net nakit pozitif mi?", "=IF(HESAP!B60<=0,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("NBD farkı hangi tarafı işaret ediyor?", "=IF(OR(HESAP!B66=\"1. ÇEYREK\",HESAP!B66=\"2. ÇEYREK\"),\"ERKEN AVANTAJLI\",\"GEÇ AVANTAJLI\")", NORMAL, "B3261E"),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Temel giriş kolonlarının dolu olma oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblKiyamet[Kıymet Türü])*AmortismanTemelKolonSayisi", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblKiyamet[Kıymet Türü])+COUNTA(tblKiyamet[Edinim Tarihi])+"
                 "COUNTA(tblKiyamet[Edinim Bedeli (TL)])+COUNTA(tblKiyamet[Amortisman Yöntemi])+"
                 "COUNTA(tblKiyamet[Amortisman Oranı])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(AmortismanToplamGiris=0,0,AmortismanDoluGiris/AmortismanToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=AmortismanKaliteIyi,\"YÜKSEK\",IF(B19>=AmortismanKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 60})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Sat, Bekle mi?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Net nakit, bugünkü değer zamanlaması ve kıymet verisi kalitesine göre "
                "üretilen satış zamanlama kararı.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Toplam Net Defter Değeri (TL)", yazi="333333")
    h(ws, 6, 2, "=HESAP!B16", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "En Yüksek Net Nakit (TL)", yazi="333333")
    h(ws, 7, 2, "=HESAP!B60", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "En İyi NBD Çeyreği", yazi="333333")
    h(ws, 8, 2, "=HESAP!B66", yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "En İyi NBD Değeri (TL)", yazi="333333")
    h(ws, 9, 2, "=HESAP!B67", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblKiyamet[Edinim Bedeli (TL)])=0,"VERİ YOK",'\
                 'IF(OR(COUNTIF(tblKiyamet[Edinim Bedeli (TL)],"<0")>0,COUNTIF(tblKiyamet[Satış Bedeli (TL)],"<0")>0),"DURDUR",'\
                 'IF(HESAP!B16<=0,"VERİ YOK",'\
                 'IF(HESAP!B60<=0,"BEKLE",'\
                 'IF(OR(HESAP!B66="1. ÇEYREK",HESAP!B66="2. ÇEYREK"),"SAT","BEKLE")))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblKiyamet[Edinim Bedeli (TL)],"<0")>0,"Negatif edinim bedeli girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(COUNTIF(tblKiyamet[Satış Bedeli (TL)],"<0")>0,"Negatif satış bedeli girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(HESAP!B16<=0,"Kıymet kaydı bulunmuyor; satış zamanlama kararı üretilemiyor.",'\
                 'IF(HESAP!B60<=0,"En iyi çeyrek net nakit pozitif değil; satışın nakit avantajı yok, satış bekletilmeli.",'\
                 'IF(OR(HESAP!B66="1. ÇEYREK",HESAP!B66="2. ÇEYREK"),"Erken çeyrek satışı bugünkü değerde en yüksek nakit akışını sağlıyor; satış yapılabilir.",'\
                 '"Yıl sonu satışı bugünkü değerde en yüksek nakit akışını sağlıyor; satış için aralık sonu beklenmeli.")))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=AmortismanAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki oran/eşiklerden türetilir; "
                     "satış sonucu bağlayıcı değildir.")
    genislik(ws, {"A": 45, "B": 60})
def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Toplam NDD (TL)", "=HESAP!B16", TL),
        (3, "En İyi Çeyrek Net Nakit (TL)", "=HESAP!B60", TL),
        (4, "En İyi NBD Çeyreği", "=HESAP!B66", None),
        (5, "En İyi NBD (TL)", "=HESAP!B67", TL),
        (6, "Toplam Yıllık Amortisman (TL)", "=HESAP!B17", TL),
        (7, "Toplam Satış Bedeli (TL)", "=HESAP!B18", TL),
        (8, "Kıymet Sayısı", "=HESAP!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Kalite Skoru", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=KONTROLLER!B19", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "En İyi NBD Çeyreği", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 8, "=HESAP!B66", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Kıymet türü bazlı grafik verisi
    h(ws, 10, 1, "Kıymet Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Edinim (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Birikmiş Amortisman (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 10, 4, "NDD (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Yıllık Amortisman (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, tur in enumerate(KIYMET_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!C{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!E{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=HESAP!F{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Kıymet kalemi bazlı NDD (kayıt sırasına göre)
    h(ws, 20, 7, "Kıymet Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 20, 8, "NDD (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 21 + i, 7, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 21 + i, 8, f"=IFERROR(INDEX(tblKiyamet[Net Defter Değeri (TL)],AYARLAR!$B${34 + i}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kıymet Türü Edinim ve Birikmiş Amortisman"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=16), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F8")
    g2 = PieChart()
    g2.title = "Net Defter Değeri Dağılımı"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=16), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Yıllık Amortisman"
    g3.add_data(Reference(ws, min_col=5, min_row=10, max_row=16), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g3.height, g3.width = 9, 16
    ws.add_chart(g3, "F21")
    g4 = LineChart()
    g4.title = "Kıymet Kalemi Bazlı NDD Trendi"
    g4.add_data(Reference(ws, min_col=8, min_row=20, max_row=32), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=7, min_row=21, max_row=32))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U21")
    # Projeksiyon
    h(ws, 34, 1, "Ortalama Yıllık Amortisman", yazi="333333")
    h(ws, 34, 2, "=ROUND(IFERROR(AVERAGE(tblKiyamet[Yıllık Amortisman (TL)]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 35, 1, "Gelecek Dönem Toplam Amortisman Tahmini", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 35, 2, '=ROUND(IFERROR(FORECAST.LINEAR(AmortismanPanelSatir12,tblKiyamet[Yıllık Amortisman (TL)],ROW(tblKiyamet[Yıllık Amortisman (TL)])-ROW(tblKiyamet[[#Headers],[Yıllık Amortisman (TL)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 36, 1, "Yıllık Amortisman Std Sapması", yazi="333333")
    h(ws, 36, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblKiyamet[Yıllık Amortisman (TL)],AmortismanPanelSatir1):INDEX(tblKiyamet[Yıllık Amortisman (TL)],AmortismanPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 37, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 37, 2, "=B34-B36*AmortismanTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 38, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 38, 2, "=B34+B36*AmortismanTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=34, max_row=38), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=35, max_row=38))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F34")
    # Çeyrek karşılaştırma
    h(ws, 40, 1, "Çeyrek", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 40, 2, "Net Nakit (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 40, 3, "NBD (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    ceyrekler = [("1. ÇEYREK", "=HESAP!B56", "=HESAP!B62"),
                 ("2. ÇEYREK", "=HESAP!B57", "=HESAP!B63"),
                 ("3. ÇEYREK", "=HESAP!B58", "=HESAP!B64"),
                 ("4. ÇEYREK", "=HESAP!B59", "=HESAP!B65")]
    for i, (ad, form, form2) in enumerate(ceyrekler, 41):
        h(ws, i, 1, ad, yazi=GRİ)
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, form2, sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = BarChart()
    g6.type = "col"
    g6.title = "Çeyrek Net Nakit Karşılaştırması"
    g6.add_data(Reference(ws, min_col=2, min_row=40, max_row=44), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=41, max_row=44))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U34")
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Çeyrek NBD Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=40, max_row=44), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=41, max_row=44))
    g7.height, g7.width = 8, 13
    ws.add_chart(g7, "F41")
    # Senaryo verisi
    h(ws, 45, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 45, 6, "Net Nakit (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 46):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-39}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g8 = LineChart()
    g8.title = "Senaryo Net Nakitleri"
    g8.add_data(Reference(ws, min_col=6, min_row=45, max_row=49), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=5, min_row=46, max_row=49))
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U41")
    baski_hazirla(ws, "A1:M50", URUN_AD)
    alt_bant(ws, 50, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 40, "B": 22, "C": 22, "D": 20, "E": 22, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Farklı satış bedeli çarpanı senaryolarında en iyi çeyrek net nakit bant aralığı; "
                "ardından hangi değişkenin net nakiti en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Satış Bedeli Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 6, 3, "En İyi Çeyrek Net Nakit (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 6, 4, "Durum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        ("İYİMSER", 18),
        ("BAZ", 19),
        ("KÖTÜMSER", 20),
        ("KRİTİK", 21),
    ]
    for i, (ad, ayar) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=AYARLAR!$B${ayar}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=IFERROR(AmortismanNetNakitEnIyi+AmortismanToplamSatis*B{i}*(1-AmortismanKurumlarVergiOran),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(C{i}>=AmortismanNetNakitEnIyi,\"AVANTAJLI\",\"DEĞERLENDİR\")", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Senaryo Bant Genişliği (İyi − Kritik)", yazi="333333", kalin=True)
    h(ws, 11, 3, "=C7-C10", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Tornado: Net Nakite Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Satış Bedeli Düşüşü Etkisi", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ABS(AmortismanToplamSatis*AYARLAR!$B$20*(1-AmortismanKurumlarVergiOran)),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Amortisman Oranı Artışı Etkisi", yazi="333333")
    h(ws, 15, 2, "=IFERROR(ABS(HESAP!B17*AYARLAR!$B$28/IF(AVERAGE(tblKiyamet[Amortisman Oranı])<>0,AVERAGE(tblKiyamet[Amortisman Oranı]),1)),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, "=IF(B14>=B15,\"Satış Bedeli\",\"Amortisman Oranı\")", yazi=DIS_REF_YEŞIL, kalin=True)
    g = BarChart()
    g.type = "bar"
    g.title = "Tornado: Net Nakite Etki"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 8, 13
    ws.add_chart(g, "E17")
    alt_bant(ws, 19, "Senaryo çarpanları ve tornado adımları AYARLAR'dan türetilir.")
    genislik(ws, {"A": 42, "B": 24, "C": 26, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "SABİT KIYMET SATIŞ ZAMANLAMA — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(AmortismanRaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",AmortismanRaporHazirlayan,\" | Sürüm: \",AmortismanDosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Toplam Net Defter Değeri (TL)", "=HESAP!B16", TL),
        ("Toplam Yıllık Amortisman (TL)", "=HESAP!B17", TL),
        ("Toplam Satış Bedeli (TL)", "=HESAP!B18", TL),
        ("En Yüksek Net Nakit (TL)", "=HESAP!B60", TL),
        ("En İyi NBD Çeyreği", "=HESAP!B66", None),
        ("En İyi NBD Değeri (TL)", "=HESAP!B67", TL),
        ("NBD Farkı (1Ç − 4Ç) (TL)", "=HESAP!B68", TL),
        ("Kıymet Sayısı", "=HESAP!B19", CATI),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "Normal yöntemde amortisman edinim bedelinin oranla çarpımıyla; azalan bakiyede VUK Md. 315 çerçevesinde normal oranın 2 katına kadar çarpanla hesaplanır.",
        "Yıl içi satış senaryolarında satış yılına isabet eden kıst amortisman AYARLAR'daki kıst oranlarıyla ayrılır (3/12, 6/12, 9/12, 12/12).",
        "Satış kârı/zararı satış bedeli ile satış anındaki net defter değeri arasındaki farktır; vergi etkisi kurumlar vergisi oranıyla hesaplanır.",
        "NBD karşılaştırması çeyrek nakit akışlarını iskonto oranıyla bugünkü değere indirger; iskonto oranı AYARLAR'dadır.",
        "Bu dosya karar destek aracıdır; amortisman ve vergi hesabı mali müşavir görüşü yerine geçmez.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=AmortismanAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    sutunlar = ["Kıymet Kodu", "Kıymet Türü", "Edinim Tarihi", "Edinim Bedeli (TL)",
                "Amortisman Yöntemi", "Amortisman Oranı", "Birikmiş Amortisman (TL)",
                "NDD (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Birikmiş Amortisman (TL)": '=IF(tblOrnekKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                                    'MIN(tblOrnekKiyamet[[#This Row],[Edinim Bedeli (TL)]],tblOrnekKiyamet[[#This Row],[Edinim Bedeli (TL)]]*'
                                    'tblOrnekKiyamet[[#This Row],[Amortisman Oranı]]*'
                                    'MAX(0,ROUNDDOWN((AmortismanRaporTarihi-tblOrnekKiyamet[[#This Row],[Edinim Tarihi]])/365,0))))',
        "NDD (TL)": '=IF(tblOrnekKiyamet[[#This Row],[Edinim Bedeli (TL)]]="",0,'
                    'MAX(0,tblOrnekKiyamet[[#This Row],[Edinim Bedeli (TL)]]-tblOrnekKiyamet[[#This Row],[Birikmiş Amortisman (TL)]]))',
        "Açıklama": '=IF(tblOrnekKiyamet[[#This Row],[Kıymet Türü]]="","",CONCATENATE("Örnek kıymet: ",tblOrnekKiyamet[[#This Row],[Kıymet Türü]]))',
    }
    tablo_ekle(ws, "tblOrnekKiyamet", "A5:I1004", sutunlar, formuller)
    ornek = [
        ("K-1001", "Bina", date(2019, 5, 15), 1500000, "Normal", 0.02),
        ("K-1002", "Makine", date(2021, 3, 10), 600000, "Azalan", 0.20),
        ("K-1003", "Makine", date(2020, 9, 1), 400000, "Normal", 0.20),
        ("K-1004", "Taşıt", date(2022, 1, 20), 350000, "Normal", 0.20),
        ("K-1005", "Demirbaş", date(2023, 4, 5), 180000, "Normal", 0.20),
        ("K-1006", "Yazılım", date(2024, 6, 15), 120000, "Azalan", 0.33),
    ]
    giris_kolonlari = [1, 2, 3, 4, 5, 6]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir):
            h(ws, i, giris_kolonlari[k], deger)
        ws.cell(row=i, column=3).number_format = TARİH
        ws.cell(row=i, column=4).number_format = TL
        ws.cell(row=i, column=6).number_format = YÜZDE
        ws.cell(row=i, column=7).number_format = TL
        ws.cell(row=i, column=8).number_format = TL
    genislik(ws, {"A": 12, "B": 18, "C": 13, "D": 18, "E": 18, "F": 15, "G": 22, "H": 18, "I": 24})
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
    h(ws, 6, 1, "Kıymet Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(KIYMET_TURLERI, 7):
        h(ws, i, 1, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Kıymet türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni tür için satır ekleyin")
    h(ws, 6, 3, "Amortisman Yöntemi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, yontem in enumerate(YONTEMLER, 7):
        h(ws, i, 3, yontem, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"C{i}", "Tanım: Amortisman yöntemi | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Normal veya Azalan seçin")
    h(ws, 6, 5, "Çeyrekler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ceyrek in enumerate(CEYREKLER, 7):
        h(ws, i, 5, ceyrek, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"E{i}", "Tanım: Çeyrek dönem | Neden önemli: Karar kapısını besler | "
                                 "Doğru kullanım: Karar çeyreği seçiminde kullanılır")
    genislik(ws, {"A": 18, "C": 22, "E": 16})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Kurumlar vergisi oranı, iskonto oranı, azalan "
                "bakiye çarpanı ve kıst oranlarını güncelleyin; değişen her değer tüm "
                "hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"])
    TARIH_D = date(2026, 8, 10)
    YURURLUK = date(2025, 1, 1)
    DOGRULAMA = date(2026, 8, 10)
    satirlar = [
        (7, "FirmaUnvan", "Örnek Firma A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "RaporTarihi", TARIH_D, "Tarih", None, None, "Amortisman ve satış hesabının dayanak tarihi.", "Kullanıcı"),
        (9, "RaporHazirlayan", "Mali İşler Direktörü", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "DosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "KurumlarVergiOran", 0.25, "Oran", 0, 1, "Kurumlar vergisi oranı; satış kârının vergi etkisinde kullanılır.", "Mevzuat (KVK Md. 32)"),
        (12, "IskontoOrani", 0.20, "Oran", 0, 1, "Çeyrek nakit akışlarını bugünkü değere indirgemedeki iskonto oranı.", "Varsayım"),
        (13, "AmortismanAzalanCarpan", 2.00, "Oran", 1, 2, "Azalan bakiyede normal oranın çarpım çarpanı (VUK en çok 2 kat).", "Mevzuat (VUK Md. 315)"),
        (14, "KistCeyrek1", 0.25, "Oran", 0, 1, "1. çeyrek satışında ayrılan yıllık amortisman oranı (3/12).", "Varsayım"),
        (15, "KistCeyrek2", 0.50, "Oran", 0, 1, "2. çeyrek satışında ayrılan yıllık amortisman oranı (6/12).", "Varsayım"),
        (16, "KistCeyrek3", 0.75, "Oran", 0, 1, "3. çeyrek satışında ayrılan yıllık amortisman oranı (9/12).", "Varsayım"),
        (17, "KistCeyrek4", 1.00, "Oran", 0, 1, "4. çeyrek satışında ayrılan yıllık amortisman oranı (12/12).", "Varsayım"),
        (18, "SenIyiSatis", 0.10, "Oran", -0.5, 0.5, "İyimser senaryo satış bedeli çarpanı.", "Varsayım"),
        (19, "SenBazSatis", 0.00, "Oran", -0.5, 0.5, "Baz senaryo satış bedeli çarpanı.", "Varsayım"),
        (20, "SenKotuSatis", -0.15, "Oran", -0.5, 0.5, "Kötümser senaryo satış bedeli çarpanı.", "Varsayım"),
        (21, "SenKritikSatis", -0.30, "Oran", -0.5, 0.5, "Kritik senaryo satış bedeli çarpanı.", "Varsayım"),
        (22, "AmortismanTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (23, "AmortismanKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (24, "AmortismanKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (25, "AmortismanNakitEsik", 100000, "TL", 0, 100000000000, "Net nakit eşiği; altında satış bekletilir.", "Varsayım"),
        (26, "AmortismanP90Oran", 0.90, "Oran", 0, 1, "NDD dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (27, "AmortismanTemelKolonSayisi", 5, "Adet", 1, 20, "Veri kalitesi skorunda temel kabul edilen zorunlu giriş kolonu sayısı.", "Varsayım"),
        (28, "AmortismanTornadoOran", 0.01, "Oran", 0, 0.1, "Amortisman oranı tornado adımı (+1 puan).", "Varsayım"),
        (29, "AmortismanTornadoVergi", 0.01, "Oran", 0, 0.1, "Kurumlar vergisi oranı tornado adımı (+1 puan).", "Varsayım"),
        (30, "AmortismanYuksekNddEsik", 2000000, "TL", 0, 100000000000, "Yüksek net defter değeri eşiği.", "Varsayım"),
        (31, "AmortismanAksiyon1_ACIKLAMA",
         "Erken çeyrek satışı bugünkü değerde avantajlı; satışı planlayın, kıst amortismanı kapatarak satış işlemini gerçekleştirin.", "Metin", None, None,
         "SAT kararında gösterilen ilk öneri.", "Varsayım"),
        (32, "AmortismanAksiyon2_ACIKLAMA",
         "Yıl sonu satışı bugünkü değerde avantajlı; satışı aralık sonuna erteleyin ve aradaki amortismanı kıst esasına göre ayırın.", "Metin", None, None,
         "BEKLE kararında gösterilen ikinci öneri.", "Varsayım"),
        (33, "AmortismanAksiyon3_ACIKLAMA",
         "Satışın nakit avantajı görünmüyor; satış bedeli tahminlerini ve kıymet portföyünü gözden geçirin, gerekirse kıymet bazlı seçim yapın.", "Metin", None, None,
         "BEKLE/DURDUR kararında gösterilen üçüncü öneri.", "Varsayım"),
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
        if 31 <= satir <= 33:
            ws.cell(row=satir, column=2).alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[satir].height = 48
        ws.row_dimensions[satir].height = 26
    for i in range(1, 13):
        satir = 34 + i
        h(ws, satir, 1, f"AmortismanPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki kıymet kalemi bazlı NDD grafiğinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki grafiği besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 34, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 22, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "KIYMET_LISTESI sayfasında sarı hücrelere kıymetlerinizi girin. Kıymet türünü ve "
         "amortisman yöntemini listeden seçin; edinim tarihi, edinim bedeli, amortisman oranı "
         "ve satış bedeli tahminini girin. Geçen yıl, birikmiş amortisman, net defter değeri "
         "ve satış kârı/zararı otomatik hesaplanır."),
        ("Çeyrek senaryoları",
         "HESAP sayfası yıl içinde dört çeyreğe göre kıst amortisman, satış kârı/zararı, "
         "kurumlar vergisi etkisi ve net nakit üretir; NBD ile bugünkü değere indirger."),
        ("Karar ve aksiyon",
         "KARAR sayfası net nakit ve NBD zamanlamasına göre SAT/BEKLE kararını gerekçesiyle "
         "üretir ve önerilen aksiyonları gösterir."),
        ("Oranlar ve senaryolar",
         "AYARLAR sayfasından kurumlar vergisi oranı, iskonto oranı, azalan bakiye çarpanı, "
         "kıst oranlar ve senaryo çarpanlarını güncelleyin; PANO ve SENARYO_DUYARLILIK anında güncellenir."),
        ("Güven ve karar kalitesi",
         "Veri kalite skoru girişlerin tamlığını ölçer. Kıymet girişi yokken karar VERİ YOK; "
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
    for deger, dolu, yazifont in (("VERİ YOK", yesil, yesil_font), ("DURDUR", kirmizi, kirmizi_font),
                                   ("SAT", yesil, yesil_font), ("BEKLE", sari, None)):
        ws.conditional_formatting.add("B11",
            CellIsRule(operator="equal", formula=[f'"{deger}"'], fill=dolu, font=yazifont))

    ws = wb["PANO"]
    for r in (4, 6):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
    for c in ("B", "C", "D", "E", "F", "G", "H"):
        for r in range(11, 17):
            ws.conditional_formatting.add(f"{c}{r}",
                CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    for c in ("B", "C", "D", "E", "F", "G", "H"):
        ws.conditional_formatting.add(f"{c}11:{c}17",
            CellIsRule(operator="greaterThan", formula=["5000000"], fill=sari))
    ws.conditional_formatting.add("B21:B32",
        CellIsRule(operator="greaterThan", formula=["1000000"], fill=koyu))

    ws = wb["HESAP"]
    for hucre in ("B16", "B28", "B36", "B44", "B52", "B60"):
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="greaterThan", formula=["AmortismanYuksekNddEsik"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B28",
        CellIsRule(operator="greaterThanOrEqual", formula=["B36"], fill=yesil))
    ws.conditional_formatting.add("B36",
        CellIsRule(operator="lessThan", formula=["B28"], fill=sari))
    for r in range(6, 12):
        ws.conditional_formatting.add(f"D{r}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

    ws = wb["KIYMET_LISTESI"]
    for c in ("D", "G"):
        ws.conditional_formatting.add(f"{c}6:{c}1004",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("H6:H1004",
        CellIsRule(operator="greaterThan", formula=["40"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("F6:F1004",
        CellIsRule(operator="greaterThan", formula=["0.5"], fill=kirmizi, font=kirmizi_font))

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
        (kIymet_listesi, "KIYMET_LISTESI"), (hesap, "HESAP"),
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
    ad_ekle(wb, "AmortismanToplamEdinim", "HESAP!$B$14")
    ad_ekle(wb, "AmortismanToplamBirikmis", "HESAP!$B$15")
    ad_ekle(wb, "AmortismanToplamNdd", "HESAP!$B$16")
    ad_ekle(wb, "AmortismanToplamYillik", "HESAP!$B$17")
    ad_ekle(wb, "AmortismanToplamSatis", "HESAP!$B$18")
    ad_ekle(wb, "AmortismanKiyametSayisi", "HESAP!$B$19")
    ad_ekle(wb, "AmortismanTurSayisi", "HESAP!$B$20")
    ad_ekle(wb, "AmortismanNetNakit1C", "HESAP!$B$28")
    ad_ekle(wb, "AmortismanNetNakit2C", "HESAP!$B$36")
    ad_ekle(wb, "AmortismanNetNakit3C", "HESAP!$B$44")
    ad_ekle(wb, "AmortismanNetNakit4C", "HESAP!$B$52")
    ad_ekle(wb, "AmortismanNetNakitEnIyi", "HESAP!$B$60")
    ad_ekle(wb, "AmortismanNbd1C", "HESAP!$B$62")
    ad_ekle(wb, "AmortismanNbd2C", "HESAP!$B$63")
    ad_ekle(wb, "AmortismanNbd3C", "HESAP!$B$64")
    ad_ekle(wb, "AmortismanNbd4C", "HESAP!$B$65")
    ad_ekle(wb, "AmortismanEnIyiCeyrek", "HESAP!$B$66")
    ad_ekle(wb, "AmortismanNbdEnIyi", "HESAP!$B$67")
    ad_ekle(wb, "AmortismanNbdFarki", "HESAP!$B$68")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "AmortismanKurumlarVergiOran": 11, "AmortismanIskontoOrani": 12, "AmortismanAzalanCarpan": 13,
        "KistCeyrek1": 14, "KistCeyrek2": 15, "KistCeyrek3": 16, "KistCeyrek4": 17,
        "SenIyiSatis": 18, "SenBazSatis": 19, "SenKotuSatis": 20, "SenKritikSatis": 21,
        "AmortismanTahminCarpan": 22, "AmortismanKaliteIyi": 23, "AmortismanKaliteOrta": 24,
        "AmortismanNakitEsik": 25, "AmortismanP90Oran": 26, "AmortismanTemelKolonSayisi": 27,
        "AmortismanTornadoOran": 28, "AmortismanTornadoVergi": 29, "AmortismanYuksekNddEsik": 30,
    }
    for i in range(1, 13):
        ayar_satirlari[f"AmortismanPanelSatir{i}"] = 34 + i
    for i in range(3):
        ayar_satirlari[f"AmortismanAksiyon{i+1}_ACIKLAMA"] = 31 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "AmortismanToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "AmortismanDoluGiris", "KONTROLLER!$B$18")
    ad_ekle(wb, "AmortismanRaporTarihi", "AYARLAR!$B$8")
    ad_ekle(wb, "AmortismanRaporHazirlayan", "AYARLAR!$B$9")
    ad_ekle(wb, "AmortismanDosyaSurumu", "AYARLAR!$B$10")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1ZamOrani": 6, "modulT1ZamOraniYorum": 6,
        "modulT2OrtalamaNdd": 7, "modulT2OrtalamaNddYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2EnYuksekKar": 9, "modulO2EnYuksekKarYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1AmortismanTahmin": 13, "modulI1AmortismanTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
        "modulI3NbdFarki": 15, "modulI3NbdFarkiYorum": 15,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeKiyametTurleri", "LISTELER!$A$7:$A$12")
    ad_ekle(wb, "ListeAmortismanYontemleri", "LISTELER!$C$7:$C$8")
    ad_ekle(wb, "ListeCeyrekler", "LISTELER!$E$7:$E$10")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "AmortismanVeSabitKiymetSatisZamanlamaStratejisti.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
