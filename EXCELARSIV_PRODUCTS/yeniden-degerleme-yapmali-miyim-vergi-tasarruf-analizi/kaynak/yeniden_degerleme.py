"""
Yeniden Değerleme Yapmalı Mıyım? (Vergi Tasarruf Analizi) — üretim betiği.
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

URUN_AD = "Yeniden Değerleme Yapmalı Mıyım? (Vergi Tasarruf Analizi)"
SURUM = "1.0.0"
RENK = "B08948"

KIYMET_TURLERI = ["Bina", "Arazi", "Makine", "Tesis", "Taşıt", "Diğer"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Kıymet bazında yeniden değerleme katsayısını işleyin; değer artışı, "
                "%2 fon vergisi ve ek amortismanın vergi tasarrufunu karşılaştırın; "
                "net faydaya göre karar verin.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Kıymet bazında değer artışı, %2 fon vergisi ve yıllık ek amortisman hesabı",
        "Ek amortismanın kalan ömür boyunca sağladığı vergi tasarrufunun bugünkü değeri",
        "Net fayda (tasarruf eksi fon) ve UYGUN / İNCELE / DURDUR karar kapısı",
        "İskonto ve fon oranına göre senaryo bant aralığı ve tornado analizi",
        "Kıymet türü bazında yoğunlaşma, anomali ve veri kalite skoru",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Yeniden değerleme yapıp yapmamayı değerlendiren mükellefler ve mali müşavirler",
        "Aktif büyüklüğünü artırmak ve ek amortisman avantajı arayan şirketler",
        "Geçici 31. madde kapsamını finansal modeliyle test etmek isteyen firmalar",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. KIYMET_LISTESI sayfasına kıymetlerinizi girin. "
                 "3. PANO ve RAPOR sayfalarından net faydayı ve kararı izleyin.", yazi="333333", kaydir=True)
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
         "KIYMET_LISTESI sayfasındaki sarı hücrelere kıymet türü, defter değeri, birikmiş "
         "amortisman, amortisman oranı, kalan faydalı ömür ve yeniden değerleme katsayısını girin."),
        ("Adım 2 — Değer artışını ve fonu izleyin",
         "HESAP sayfası değer artışını, %2 fon vergisini ve yıllık ek amortismanı hesaplar."),
        ("Adım 3 — Vergi tasarrufunu karşılaştırın",
         "Ek amortismanın kalan ömür boyunca sağladığı tasarrufun bugünkü değeri fon maliyetiyle "
         "karşılaştırılır; KARAR sayfası UYGUN/İNCELE/DURDUR üretir."),
        ("Adım 4 — Oranları ayarlayın",
         "Fon oranı, kurumlar vergisi oranı ve iskonto oranı AYARLAR sayfasından değiştirilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Amortismanı defter değerinden büyük girmek (net defter değeri negatif olur)",
        "Araziyi amortismana tabi kıymet gibi işlemek (arazide ek amortisman yoktur)",
        "Katsayıyı gerçekleşmemiş oranlarla doldurmak (oran güncel Yİ-ÜFE katsayısına göre girilir)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def kiyamet_listesi(ws):
    sayfa_hazirla(ws, "KIYMET_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kıymet Listesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Her satır bir duran varlık kalemidir; "
                "değer artışı, fon vergisi ve tasarruf otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kayıt Sırası", "Kıymet Türü", "Defter Değeri (Brüt)", "Birikmiş Amortisman",
                "Net Defter Değeri", "Amortisman Oranı", "Kalan Faydalı Ömür",
                "Yeniden Değerleme Katsayısı", "Değerlenmiş Net Değer", "Değer Artışı",
                "Fon Vergisi", "Yıllık Ek Amortisman", "Yıllık Vergi Tasarrufu",
                "Tasarruf Bugünkü Değeri", "Net Fayda", "Öneri", "Kayıt Açıklaması"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Net Defter Değeri": "=IF(tblKiyamet[[#This Row],[Defter Değeri (Brüt)]]=0,0,"\
                             "tblKiyamet[[#This Row],[Defter Değeri (Brüt)]]-tblKiyamet[[#This Row],[Birikmiş Amortisman]])",
        "Değerlenmiş Net Değer": "=IF(tblKiyamet[[#This Row],[Net Defter Değeri]]=0,0,"\
                                 "tblKiyamet[[#This Row],[Net Defter Değeri]]*tblKiyamet[[#This Row],[Yeniden Değerleme Katsayısı]])",
        "Değer Artışı": "=IF(tblKiyamet[[#This Row],[Değerlenmiş Net Değer]]=0,0,"\
                        "tblKiyamet[[#This Row],[Değerlenmiş Net Değer]]-tblKiyamet[[#This Row],[Net Defter Değeri]])",
        "Fon Vergisi": "=IF(tblKiyamet[[#This Row],[Değer Artışı]]=0,0,"\
                       "tblKiyamet[[#This Row],[Değer Artışı]]*YenidegerFonOran)",
        "Yıllık Ek Amortisman": "=IF(tblKiyamet[[#This Row],[Kalan Faydalı Ömür]]=0,0,"\
                                "tblKiyamet[[#This Row],[Değer Artışı]]/tblKiyamet[[#This Row],[Kalan Faydalı Ömür]])",
        "Yıllık Vergi Tasarrufu": "=IF(tblKiyamet[[#This Row],[Yıllık Ek Amortisman]]=0,0,"\
                                   "tblKiyamet[[#This Row],[Yıllık Ek Amortisman]]*YenidegerKurumlarVergiOran)",
        "Tasarruf Bugünkü Değeri": "=IFERROR(IF(YenidegerIskontoOrani=0,"\
                                   "tblKiyamet[[#This Row],[Yıllık Vergi Tasarrufu]]*tblKiyamet[[#This Row],[Kalan Faydalı Ömür]],"\
                                   "tblKiyamet[[#This Row],[Yıllık Vergi Tasarrufu]]*(1-1/(1+YenidegerIskontoOrani)^tblKiyamet[[#This Row],[Kalan Faydalı Ömür]])/YenidegerIskontoOrani),0)",
        "Net Fayda": "=tblKiyamet[[#This Row],[Tasarruf Bugünkü Değeri]]-tblKiyamet[[#This Row],[Fon Vergisi]]",
        "Öneri": "=IF(tblKiyamet[[#This Row],[Değer Artışı]]<=0,\"DEĞERLENDİR\","\
                 "IF(tblKiyamet[[#This Row],[Net Fayda]]>0,\"YENİDEN DEĞERLE\",\"DEĞERLENDİR\"))",
        "Kayıt Açıklaması": "=IF(tblKiyamet[[#This Row],[Kıymet Türü]]=\"\",\"\","\
                            "CONCATENATE(\"Kıymet: \",tblKiyamet[[#This Row],[Kıymet Türü]],\" | Değer Artışı: \",TEXT(tblKiyamet[[#This Row],[Değer Artışı]],\"#,##0\")))",
    }
    tablo_ekle(ws, "tblKiyamet", "A5:Q1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeKiyametTurleri", "B6:B1004",
              baslik="Kıymet Türü", mesaj="Listeden bir kıymet türü seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "whole", "1", "A6:A1004",
              baslik="Kayıt Sırası", mesaj="1 ile 99999 arasında sıra numarası girin.",
              hata_baslik="Geçersiz sıra", hata_mesaj="Sıra 1 ile 99999 arasında olmalı.",
              isaret="between", f2="99999")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Defter Değeri", mesaj="KDV hariç brüt defter değerini TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Birikmiş Amortisman", mesaj="Bugüne kadar ayrılan amortismanı TL girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Amortisman Oranı", mesaj="Amortisman oranını 0-1 arasında girin (örn. %10 için 0,10).",
              hata_baslik="Geçersiz oran", hata_mesaj="Oran 0 ile 1 arasında olmalı.",
              isaret="between", f2="1")
    dogrulama(ws, "decimal", "0", "G6:G1004",
              baslik="Kalan Faydalı Ömür", mesaj="Kalan faydalı ömrü yıl olarak girin.",
              hata_baslik="Geçersiz ömür", hata_mesaj="Ömür 0 ile 100 yıl arasında olmalı.",
              isaret="between", f2="100")
    dogrulama(ws, "decimal", "0", "H6:H1004",
              baslik="Yeniden Değerleme Katsayısı", mesaj="Dönem yeniden değerleme katsayısını girin.",
              hata_baslik="Geçersiz katsayı", hata_mesaj="Katsayı 0 ile 100 arasında olmalı.",
              isaret="between", f2="100")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = CATI
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=6).number_format = YÜZDE
        ws.cell(row=satir, column=7).number_format = CATI
        ws.cell(row=satir, column=8).number_format = "0.00"
    ornek_kiymetler = [
        (1, "Bina", 500000, 200000, 0.02, 10, 2.5),
        (2, "Bina", 300000, 100000, 0.02, 8, 2.2),
        (3, "Makine", 250000, 120000, 0.10, 7, 3.1),
        (4, "Makine", 400000, 220000, 0.10, 5, 2.8),
        (5, "Tesis", 1500000, 600000, 0.04, 15, 2.3),
        (6, "Taşıt", 120000, 80000, 0.20, 4, 1.9),
        (7, "Arazi", 1000000, 0, 0.00, 0, 3.0),
        (8, "Bina", 250000, 90000, 0.02, 12, 2.6),
        (9, "Makine", 180000, 70000, 0.10, 6, 3.2),
        (10, "Tesis", 800000, 300000, 0.04, 20, 2.4),
    ]
    giris_kolonlari = [1, 2, 3, 4, 6, 7, 8]
    for i, satir in enumerate(ornek_kiymetler, 6):
        for k, deger in enumerate(satir):
            ws.cell(row=i, column=giris_kolonlari[k]).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 6, 7, 8])
    sabitle(ws, "A6")
    genislik(ws, {"A": 11, "B": 15, "C": 18, "D": 18, "E": 18, "F": 13, "G": 14,
                  "H": 15, "I": 18, "J": 18, "K": 15, "L": 18, "M": 18, "N": 20,
                  "O": 16, "P": 20, "Q": 26})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kıymet Türü Bazlı Yeniden Değerleme Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet kıymet türlerine göre listeden otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kıymet Türü", "Net Defter Değeri", "Değer Artışı", "Fon Vergisi",
                "Yıllık Ek Amortisman", "Yıllık Vergi Tasarrufu", "Tasarruf Bugünkü Değeri",
                "Net Fayda", "Öneri"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(KIYMET_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=SUMIFS(tblKiyamet[Net Defter Değeri],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblKiyamet[Değer Artışı],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblKiyamet[Fon Vergisi],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblKiyamet[Yıllık Ek Amortisman],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=SUMIFS(tblKiyamet[Yıllık Vergi Tasarrufu],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f'=SUMIFS(tblKiyamet[Tasarruf Bugünkü Değeri],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 8, f'=SUMIFS(tblKiyamet[Net Fayda],tblKiyamet[Kıymet Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 9, f'=IF(B{r}<=0,"VERİ YOK",IF(H{r}>0,"YENİDEN DEĞERLE","DEĞERLENDİR"))',
          yazi=DIS_REF_YEŞIL)
        for k in range(2, 10):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: {tur} yeniden değerleme metriği | "
                       f"Neden önemli: Değer artışı ve net faydayı gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    t = 6 + len(KIYMET_TURLERI)
    h(ws, t, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    for k in range(2, 9):
        h(ws, t, k, f"=SUM({chr(64+k)}6:{chr(64+k)}{t-1})", sayi=TL,
          yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t, 9, f"=IF(H{t}>0,\"YENİDEN DEĞERLE\",\"DEĞERLENDİR\")", yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 1, 1, "Toplam Net Defter Değeri (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 1, 2, f"=B{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 2, 1, "Net Fayda (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 2, 2, f"=H{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 3, 1, "Fon Vergisi Yükü (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 3, 2, f"=D{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 4, 1, "Toplam Değer Artışı (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 4, 2, f"=C{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 5, 1, "Yıllık Vergi Tasarrufu (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 5, 2, f"=F{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 6, 1, "Tasarruf Bugünkü Değeri (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 6, 2, f"=G{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 7, 1, "Kalan Ömür Ortalaması (Yıl)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 7, 2, "=IFERROR(AVERAGE(tblKiyamet[Kalan Faydalı Ömür]),0)", sayi="0.0",
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 8, 1, "Kıymet Adedi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 8, 2, "=COUNTA(tblKiyamet[Kıymet Türü])", sayi=CATI,
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 9, 1, "Net Fayda Verimliliği (NBD/Fon)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 9, 2, "=IFERROR(H12/D12,0)", sayi="0.00", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, t + 11, "Hesap satırları yalnızca kıymet listesinden beslenir; korumalıdır.")
    genislik(ws, {"A": 36, "B": 16, "C": 16, "D": 16, "E": 18, "F": 18, "G": 20, "H": 16, "I": 20})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Ortalama Değer Artışı Oranı",
         "=IFERROR(IF(AVERAGE(tblKiyamet[Net Defter Değeri])=0,0,"\
         "AVERAGE(tblKiyamet[Değer Artışı])/AVERAGE(tblKiyamet[Net Defter Değeri])),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama değer artışı oranı ",TEXT(C6,"0.0%")," — katsayı uygulamasının ortalama büyütme gücünü gösterir")'),
        ("T2", "Ortalama Yeniden Değerleme Katsayısı",
         "=IFERROR(AVERAGE(tblKiyamet[Yeniden Değerleme Katsayısı]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama yeniden değerleme katsayısı ",TEXT(C7,"0.00")," — katsayı büyüdükçe fon yükü de büyür")'),
        ("O1", "Net Fayda Sapması (MAD)",
         "=IFERROR(AVEDEV(tblKiyamet[Net Fayda]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Net fayda ortalama mutlak sapması ",TEXT(C8,"₺ #,##0")," TL — sapma yüksekse kıymet portföyü heterojendir")'),
        ("O2", "En Yüksek Net Fayda",
         "=IFERROR(MAX(tblKiyamet[Net Fayda]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek net faydalı kıymet ",TEXT(C9,"₺ #,##0")," TL — bu kıymet yeniden değerlemeye en uygun adaydır")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C10",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kritik senaryo net fayda farkı ",TEXT(C10,"₺ #,##0")," TL — iskonto belirsizliğinin etkisini gösterir")'),
        ("O6", "Değer Artışı Yoğunlaşması (HHI)",
         "=IFERROR(IF(SUM(HESAP!$C$6:$C$11)=0,0,"\
         "SUMPRODUCT((HESAP!$C$6:$C$11/SUM(HESAP!$C$6:$C$11))^2)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Değer artışı yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça fayda tek türden gelir")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=YenidegerKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Net Fayda Tahmini",
         "=PANO!B30",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem net fayda tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Net Fayda Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblKiyamet[Net Fayda],YenidegerP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kıymet net faydalarının %90 ı ",TEXT(C14,"₺ #,##0")," altında gerçekleşir — üst sınır portföy seçiminde kullanılabilir")'),
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
        ("Boş kıymet türü satırı var mı?", "=IF(COUNTBLANK(tblKiyamet[Kıymet Türü])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, "B3261E"),
        ("Boş defter değeri var mı?", "=IF(COUNTBLANK(tblKiyamet[Defter Değeri (Brüt)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, "B3261E"),
        ("Negatif defter değeri var mı?", "=IF(COUNTIF(tblKiyamet[Defter Değeri (Brüt)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Negatif amortisman var mı?", "=IF(COUNTIF(tblKiyamet[Birikmiş Amortisman],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Katsayı 1'in altında mı?", "=IF(COUNTIF(tblKiyamet[Yeniden Değerleme Katsayısı],\"<1\")=0,\"TEMİZ\",\"KATSAYI <1 VAR\")", NORMAL, "B3261E"),
        ("Kalan ömür boş mu?", "=IF(COUNTBLANK(tblKiyamet[Kalan Faydalı Ömür])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, "B3261E"),
        ("Fon yükü eşiği aşılıyor mu?", "=IF(HESAP!D12>YenidegerFonEsik,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("Net fayda eşiği aşılıyor mu?", "=IF(HESAP!H12>YenidegerNbdEsik,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Temel giriş kolonlarının dolu olma oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblKiyamet[Kıymet Türü])*YenidegerTemelKolonSayisi", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblKiyamet[Kıymet Türü])+COUNTA(tblKiyamet[Defter Değeri (Brüt)])+"
                 "COUNTA(tblKiyamet[Birikmiş Amortisman])+COUNTA(tblKiyamet[Kalan Faydalı Ömür])+"
                 "COUNTA(tblKiyamet[Yeniden Değerleme Katsayısı])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(YenidegerToplamGiris=0,0,YenidegerDoluGiris/YenidegerToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=YenidegerKaliteIyi,\"YÜKSEK\",IF(B19>=YenidegerKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 60})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Yeniden Değerleme Vergi Tasarrufu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Net fayda (tasarruf eksi fon), değer artışı ve veri kalitesine göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Toplam Değer Artışı (TL)", yazi="333333")
    h(ws, 6, 2, "=HESAP!C12", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Fon Vergisi Yükü (TL)", yazi="333333")
    h(ws, 7, 2, "=HESAP!D12", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Yıllık Vergi Tasarrufu (TL)", yazi="333333")
    h(ws, 8, 2, "=HESAP!F12", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Net Fayda (TL)", yazi="333333")
    h(ws, 9, 2, "=HESAP!H12", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblKiyamet[Kıymet Türü])=0,"VERİ YOK",IF(COUNTIF(tblKiyamet[Defter Değeri (Brüt)],"<0")>0,"DURDUR",IF(COUNTIF(tblKiyamet[Birikmiş Amortisman],"<0")>0,"DURDUR",IF(SUM(tblKiyamet[Net Fayda])>YenidegerNbdEsik,"UYGUN",IF(SUM(tblKiyamet[Değer Artışı])>YenidegerArtisEsik,"İNCELE","DURDUR")))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblKiyamet[Defter Değeri (Brüt)],"<0")>0,"Negatif defter değeri girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(COUNTIF(tblKiyamet[Birikmiş Amortisman],"<0")>0,"Negatif amortisman girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(SUM(tblKiyamet[Net Fayda])>YenidegerNbdEsik,"Yeniden değerlemenin vergi tasarrufu bugünkü değeri fon maliyetinin üzerinde; yeniden değerleme önerilir.",'\
                 'IF(SUM(tblKiyamet[Değer Artışı])>YenidegerArtisEsik,"Değer artışı eşiğin üzerinde ancak net fayda pozitif değil; kalan ömür, iskonto oranı ve fon oranı yeniden incelenmeli.",'\
                 '"Karar kapısı uygun değil; yeniden değerleme net fayda üretmiyor."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=YenidegerAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki oran/eşiklerden türetilir; "
                     "mali müşavir görüşü değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Toplam Net Defter Değeri (TL)", "=HESAP!B12", TL),
        (3, "Toplam Değer Artışı (TL)", "=HESAP!C12", TL),
        (4, "Fon Vergisi Yükü (TL)", "=HESAP!D12", TL),
        (5, "Yıllık Vergi Tasarrufu (TL)", "=HESAP!F12", TL),
        (6, "Tasarruf Bugünkü Değeri (TL)", "=HESAP!G12", TL),
        (7, "Net Fayda (TL)", "=HESAP!H12", TL),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Kalite Skoru", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=KONTROLLER!B19", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "Kıymet Adedi", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 8, "=HESAP!B20", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — kıymet türü bazlı
    h(ws, 10, 1, "Kıymet Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Değer Artışı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Net Fayda (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Fon Vergisi (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Yıllık Tasarruf (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, tur in enumerate(KIYMET_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!C{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!H{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!D{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=HESAP!F{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Kıymet bazlı net fayda (kayıt sırasına göre)
    h(ws, 10, 7, "Kıymet Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 8, "Net Fayda (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(10):
        h(ws, 11 + i, 7, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 11 + i, 8, f"=IFERROR(INDEX(tblKiyamet[Net Fayda],AYARLAR!$B${31 + i}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Değer Artışı ve Net Fayda"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=16), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F8")
    g2 = PieChart()
    g2.title = "Değer Artışı Dağılımı"
    g2.add_data(Reference(ws, min_col=2, min_row=10, max_row=16), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Fon Vergisi ve Yıllık Tasarruf"
    g3.add_data(Reference(ws, min_col=4, min_row=10, max_col=5, max_row=16), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g3.height, g3.width = 9, 16
    ws.add_chart(g3, "F21")
    g4 = LineChart()
    g4.title = "Kıymet Bazlı Net Fayda Trendi"
    g4.add_data(Reference(ws, min_col=8, min_row=10, max_row=20), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=7, min_row=11, max_row=20))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U21")
    # Zaman serisi projeksiyonu
    h(ws, 30, 1, "Gelecek Dönem Net Fayda Tahmini", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, '=ROUND(IFERROR(FORECAST.LINEAR(YenidegerPanelSatir12,tblKiyamet[Net Fayda],ROW(tblKiyamet[Net Fayda])-ROW(tblKiyamet[[#Headers],[Net Fayda]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 31, 1, "Ortalama Net Fayda", yazi="333333")
    h(ws, 31, 2, "=ROUND(IFERROR(AVERAGE(tblKiyamet[Net Fayda]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 32, 1, "Son 10 Kıymet Net Fayda Standart Sapması", yazi="333333")
    h(ws, 32, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblKiyamet[Net Fayda],YenidegerPanelSatir1):INDEX(tblKiyamet[Net Fayda],YenidegerPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 33, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 33, 2, "=B31-B32*YenidegerTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 34, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 34, 2, "=B31+B32*YenidegerTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=30, max_row=34), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=31, max_row=34))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F34")
    g6 = BarChart()
    g6.type = "col"
    g6.title = "Fon Vergisi Yapısı"
    g6.add_data(Reference(ws, min_col=4, min_row=10, max_row=16), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U34")
    # Senaryo karşılaştırma verisi
    h(ws, 37, 2, "Senaryo Net Fayda", kalin=True, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 38):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i-31}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo Net Fayda Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=37, max_row=41), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=38, max_row=41))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F41")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Yıllık Vergi Tasarrufu Yapısı"
    g8.add_data(Reference(ws, min_col=5, min_row=10, max_row=16), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U41")
    baski_hazirla(ws, "A1:M48", URUN_AD)
    alt_bant(ws, 48, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 40, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Farklı iskonto oranlarında net fayda bant aralığı; ardından hangi değişkenin "
                "sonucu en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "İskonto Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Net Fayda (NBD-Fon)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "Durum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        ("İYİMSER", 14),
        ("BAZ", 15),
        ("KÖTÜMSER", 16),
        ("KRİTİK", 17),
    ]
    for i, (ad, ayar) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=AYARLAR!$B${ayar}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "=IFERROR(YenidegerYillikTasarruf*(1-1/(1+AYARLAR!$B$" + str(ayar) + ")^YenidegerKalanOmurOrtalama)/AYARLAR!$B$" + str(ayar) + "-YenidegerFonToplam,0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(C{i}>0,\"ÖNERİLİR\",\"DEĞERLENDİR\")", yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Tornado: Net Faydaya Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 13, 1, "İskonto Oranı +%1 Etkisi", yazi="333333")
    h(ws, 13, 2, "=IFERROR(ABS(YenidegerYillikTasarruf*(1-1/(1+AYARLAR!$B$15+AYARLAR!$B$21)^YenidegerKalanOmurOrtalama)/(AYARLAR!$B$15+AYARLAR!$B$21)-C8),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Fon Oranı +%1 Etkisi", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ABS(YenidegerFonToplam*AYARLAR!$B$22/AYARLAR!$B$11),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, "=IF(B13>=B14,\"İskonto Oranı\",\"Fon Oranı\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 18, "Senaryo iskonto oranları ve tornado çarpanları AYARLAR'dan türetilir.")
    genislik(ws, {"A": 42, "B": 30, "C": 22, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "YENİDEN DEĞERLEME — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Toplam Değer Artışı (TL)", "=HESAP!C12", TL),
        ("Fon Vergisi Yükü (TL)", "=HESAP!D12", TL),
        ("Yıllık Vergi Tasarrufu (TL)", "=HESAP!F12", TL),
        ("Tasarruf Bugünkü Değeri (TL)", "=HESAP!G12", TL),
        ("Net Fayda (TL)", "=HESAP!H12", TL),
        ("Net Fayda Verimliliği (NBD/Fon)", "=HESAP!B21", "0.00"),
        ("Kıymet Adedi", "=HESAP!B20", CATI),
        ("Ortalama Değer Artışı Oranı", "=ANALITIK_MOTOR!C6", YÜZDE),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "Değer artışı, değerlenmiş net değer eksi net defter değeri olarak hesaplanır.",
        "Fon vergisi, değer artışı üzerinden VUK Geçici 31 kapsamında %2 olarak alınır (oran AYARLAR'dan değişir).",
        "Vergi tasarrufu, yıllık ek amortismanın kurumlar vergisi oranı ile çarpımıdır; iskonto oranıyla bugünkü değere indirgenir.",
        "Arazi gibi amortismana tabi olmayan kıymetlerde ek amortisman yoktur; net fayda fon yükü kadar negatif kalır.",
        "Bu dosya karar destek aracıdır; mali müşavir veya vergi danışmanı görüşü yerine geçmez.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=YenidegerAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    sutunlar = ["Kayıt Sırası", "Kıymet Türü", "Defter Değeri (Brüt)", "Birikmiş Amortisman",
                "Net Defter Değeri", "Yeniden Değerleme Katsayısı", "Değer Artışı", "Not"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Net Defter Değeri": "=IF(tblOrnek[[#This Row],[Defter Değeri (Brüt)]]=0,0,"\
                             "tblOrnek[[#This Row],[Defter Değeri (Brüt)]]-tblOrnek[[#This Row],[Birikmiş Amortisman]])",
        "Değer Artışı": "=IF(tblOrnek[[#This Row],[Net Defter Değeri]]=0,0,"\
                        "tblOrnek[[#This Row],[Net Defter Değeri]]*(tblOrnek[[#This Row],[Yeniden Değerleme Katsayısı]]-1))",
        "Not": "=IF(tblOrnek[[#This Row],[Kıymet Türü]]=\"\",\"\",CONCATENATE(\"Örnek kıymet: \",tblOrnek[[#This Row],[Kıymet Türü]]))",
    }
    tablo_ekle(ws, "tblOrnek", "A5:H1004", sutunlar, formuller)
    ornek = [
        (1, "Bina", 500000, 200000, 2.5),
        (2, "Bina", 300000, 100000, 2.2),
        (3, "Makine", 250000, 120000, 3.1),
        (4, "Makine", 400000, 220000, 2.8),
        (5, "Tesis", 1500000, 600000, 2.3),
        (6, "Taşıt", 120000, 80000, 1.9),
    ]
    giris_kolonlari = [1, 2, 3, 4, 6]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir):
            h(ws, i, giris_kolonlari[k], deger)
        ws.cell(row=i, column=1).number_format = CATI
        ws.cell(row=i, column=3).number_format = TL
        ws.cell(row=i, column=4).number_format = TL
        ws.cell(row=i, column=5).number_format = TL
        ws.cell(row=i, column=6).number_format = "0.00"
        ws.cell(row=i, column=7).number_format = TL
    genislik(ws, {"A": 11, "B": 15, "C": 18, "D": 18, "E": 18, "F": 15, "G": 18, "H": 22})
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
    genislik(ws, {"A": 18})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Fon oranı, kurumlar vergisi oranı ve iskonto "
                "oranını güncelleyin; değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"], basla=1)
    YURURLUK = date(2026, 1, 1)
    DOGRULAMA = date(2026, 8, 10)
    ayarlar = [
        (7, "FirmaUnvan", "Örnek Firma A.Ş.", "Metin", "", "", "Rapor ve kapakta görünen firma adı.", "Kullanıcı"),
        (8, "RaporTarihi", date(2026, 8, 10), "Tarih", "", "", "Tüm hesapların ve kararın alındığı tarih.", "Kullanıcı"),
        (9, "RaporHazirlayan", "Mali İşler Direktörü", "Metin", "", "", "Raporu hazırlayan kişi/rol.", "Kullanıcı"),
        (10, "DosyaSurumu", "1.0.0", "Metin", "", "", "Ürün sürüm bilgisi; rapor alt bilgisinde görünür.", "Kullanıcı"),
        (11, "YenidegerFonOran", 0.02, "Oran", 0, 0.1, "Değer artışı üzerinden alınan fon vergisi oranı (%2).", "Mevzuat (VUK Geçici 31)"),
        (12, "YenidegerKurumlarVergiOran", 0.25, "Oran", 0, 1, "Kurumlar vergisi oranı (%25); tasarruf hesabında kullanılır.", "Mevzuat (KVK Md. 32)"),
        (13, "YenidegerIskontoOrani", 0.15, "Oran", 0, 1, "Tasarrufları bugünkü değere indirgemede kullanılan iskonto oranı.", "Varsayım"),
        (14, "YenidegerIyiIskonto", 0.10, "Oran", 0, 1, "İyimser senaryo iskonto oranı.", "Varsayım"),
        (15, "YenidegerBazIskonto", 0.15, "Oran", 0, 1, "Baz senaryo iskonto oranı.", "Varsayım"),
        (16, "YenidegerKotuIskonto", 0.22, "Oran", 0, 1, "Kötümser senaryo iskonto oranı.", "Varsayım"),
        (17, "YenidegerKritikIskonto", 0.30, "Oran", 0, 1, "Kritik senaryo iskonto oranı.", "Varsayım"),
        (18, "YenidegerNbdEsik", 100000, "TL", 0, 10000000000, "Net fayda eşiği; üzerinde karar UYGUN olur.", "Varsayım"),
        (19, "YenidegerArtisEsik", 50000, "TL", 0, 10000000000, "Değer artışı eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (20, "YenidegerFonEsik", 200000, "TL", 0, 10000000000, "Fon yükü eşiği; üzerinde kontroller AŞIYOR üretir.", "Varsayım"),
        (21, "YenidegerTornadoIskonto", 0.01, "Oran", 0, 0.1, "İskonto oranının %1 artışının net faydaya etkisi.", "Varsayım"),
        (22, "YenidegerTornadoFon", 0.01, "Oran", 0, 0.1, "Fon oranının %1 artışının net faydaya etkisi.", "Varsayım"),
        (23, "YenidegerTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (24, "YenidegerKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (25, "YenidegerKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (26, "YenidegerAksiyon1_ACIKLAMA",
         "Net fayda eşiğin üzerinde: yeniden değerleme yapılabilir; değer artışını dönem beyanına işleyin ve ek amortismanı kalan ömür boyunca ayırın.", "Metin", "", "",
         "UYGUN kararında gösterilen ilk öneri.", "Varsayım"),
        (27, "YenidegerAksiyon2_ACIKLAMA",
         "Net fayda pozitif değil: kalan faydalı ömrü, iskonto oranını ve fon oranını mali müşavirinizle yeniden gözden geçirin; kıymet bazlı seçim yapın.", "Metin", "", "",
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (28, "YenidegerAksiyon3_ACIKLAMA",
         "Yeniden değerleme bu portföyde vergi tasarrufu üretmiyor; katsayı ve oran değişimlerini izleyin, uygun dönemde yeniden değerlendirin.", "Metin", "", "",
         "DURDUR kararında gösterilen üçüncü öneri.", "Varsayım"),
        (29, "YenidegerP90Oran", 0.9, "Oran", 0.5, 1, "Net fayda dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (30, "YenidegerTemelKolonSayisi", 5, "Adet", 1, 20, "Veri kalitesi skorunda temel kabul edilen zorunlu giriş kolonu sayısı.", "Varsayım"),
    ]
    for satir, anahtar, deger, birim, alt, ust, aciklama, kaynak in ayarlar:
        h(ws, satir, 1, anahtar, kalin=True, zemin=ACIK_GRI, boyut=9, kaydir=True)
        hb = h(ws, satir, 2, deger, yazi=GIRIS_YAZI, zemin=GIRIS_SARI, kalin=True,
               kaydir=(isinstance(deger, str) and len(deger) > 45))
        h(ws, satir, 3, birim, boyut=9, yazi=GRİ)
        h(ws, satir, 4, alt if alt != "" else None, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 5, ust if ust != "" else None, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 6, aciklama, boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, kaynak, boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        if isinstance(deger, date):
            hb.number_format = TARİH
        elif isinstance(deger, (int, float)):
            hb.number_format = TL if birim == "TL" else YÜZDE if birim == "Oran" else CATI
        if alt != "":
            dv = DataValidation(type="decimal", operator="between", formula1=str(alt),
                                formula2=str(ust), allow_blank=False,
                                promptTitle=anahtar, prompt=aciklama,
                                showErrorMessage=True,
                                errorTitle="Değer aralık dışı", error=aciklama)
            ws.add_data_validation(dv)
            dv.add(f"B{satir}")
        ws.row_dimensions[satir].height = 26
    for i in range(1, 13):
        satir = 31 + i
        h(ws, satir, 1, f"YenidegerPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki kıymet bazlı net fayda grafiğinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki grafiği besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 32, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 20, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "KIYMET_LISTESI sayfasında sarı hücrelere kıymetlerinizi girin. Kıymet türünü listeden "
         "seçin; defter değeri, birikmiş amortisman, amortisman oranı, kalan faydalı ömür ve "
         "yeniden değerleme katsayısını girin."),
        ("Hesap katmanı",
         "HESAP sayfası kıymet türü bazında değer artışını, %2 fon vergisini, yıllık ek amortismanı "
         "ve vergi tasarrufunun bugünkü değerini otomatik özetler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki fon oranı, kurumlar vergisi oranı ve iskonto oranını mali "
         "müşavirinize göre güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; mali müşavir ve vergi danışmanı görüşü niteliği taşımaz.",
      yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "İskonto Oranı ve Senaryo Oranları (iyimser-baz-kötümser-kritik)",
        "Net Fayda ve Değer Artışı Eşikleri",
        "Fon Yükü Eşiği",
        "Tornado İskonto ve Fon Etkisi Oranları",
        "Tahmin Çarpanı (standart sapma katı)",
        "Kalite Skoru İyi ve Orta Eşikleri",
        "P90 Yüzdelik Oranı",
        "Aksiyon Önerisi Açıklamaları",
    ]
    for j, v in enumerate(varsayimlar, 15):
        h(ws, j, 1, "• " + v, yazi="333333", boyut=9, kaydir=True)
    h(ws, 24, 1, "Bu dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. "
                "Makro ve dış bağlantı yoktur.", yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells(start_row=24, start_column=2, end_row=24, end_column=20)
    genislik(ws, {"A": 45, "B": 60})


def kosullu_bicimlendirme(wb):
    yesil, amber, kirmizi = "1F7A4D", "B7791F", "B3261E"

    def f(renk):
        return Font(name=FONT, color=renk)

    def d(renk):
        return PatternFill("solid", start_color=renk, end_color=renk)

    ws = wb["KIYMET_LISTESI"]
    ws.conditional_formatting.add("C6:C1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("J6:J1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("H6:H1004", CellIsRule(operator="lessThan", formula=["1"], font=f(amber)))
    ws.conditional_formatting.add("O6:O1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("O6:O1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("K6:K1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))
    ws.conditional_formatting.add("P6:P1004", CellIsRule(operator="equal", formula=['"YENİDEN DEĞERLE"'], font=f(yesil)))
    ws.conditional_formatting.add("P6:P1004", CellIsRule(operator="equal", formula=['"DEĞERLENDİR"'], font=f(amber)))

    ws = wb["HESAP"]
    for harf in ("B", "C", "D", "E", "F", "G", "H"):
        ws.conditional_formatting.add(f"{harf}6:{harf}11",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("C6:C11", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("H6:H11", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("H6:H11", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D11", CellIsRule(operator="greaterThan", formula=["200000"], font=f(amber)))
    ws.conditional_formatting.add("I6:I11", CellIsRule(operator="equal", formula=['"YENİDEN DEĞERLE"'], font=f(yesil)))
    ws.conditional_formatting.add("I6:I11", CellIsRule(operator="equal", formula=['"DEĞERLENDİR"'], font=f(amber)))
    ws.conditional_formatting.add("D12:D12", CellIsRule(operator="greaterThan", formula=["200000"], font=f(amber)))
    ws.conditional_formatting.add("H12:H12", CellIsRule(operator="greaterThan", formula=["100000"], font=f(yesil)))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "AŞIYOR", "KATSAYI <1 VAR"):
        ws.conditional_formatting.add("B6:B13", CellIsRule(operator="equal", formula=[f'"{deger}"'],
                                     font=f(kirmizi), fill=d("FDE9E9")))
    for deger in ("TEMİZ", "NORMAL"):
        ws.conditional_formatting.add("B6:B13", CellIsRule(operator="equal", formula=[f'"{deger}"'],
                                     font=f(yesil)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="lessThan", formula=["70"], font=f(kirmizi)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="between", formula=["70", "90"], font=f(amber)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="greaterThanOrEqual", formula=["90"], font=f(yesil)))

    ws = wb["KARAR"]
    ws.conditional_formatting.add("B11:B11", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B11:B11", CellIsRule(operator="equal", formula=['"İNCELE"'],
                                 font=f(amber), fill=d("FFF2CC")))
    ws.conditional_formatting.add("B11:B11", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B7:B7", CellIsRule(operator="greaterThan", formula=["200000"], font=f(amber)))
    ws.conditional_formatting.add("B9:B9", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["PANO"]
    for harf in ("B", "C", "D", "E", "F", "G"):
        ws.conditional_formatting.add(f"{harf}4:{harf}4",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("C4:C4", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("G4:G4", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:C16", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B11:C16", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("D11:E16", CellIsRule(operator="greaterThan", formula=["200000"], font=f(amber)))
    ws.conditional_formatting.add("H11:H20", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("H11:H20", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("C38:C41", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B33:B33", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C10", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B7:B10", CellIsRule(operator="greaterThan", formula=["0.2"], font=f(kirmizi)))
    ws.conditional_formatting.add("B13:B14", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["ORNEK_VERI"]
    ws.conditional_formatting.add("C6:C1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()
    kapak(ws)
    sayfalar = [
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("KIYMET_LISTESI", kiyamet_listesi),
        ("HESAP", hesap),
        ("ANALITIK_MOTOR", analitik_motor),
        ("KONTROLLER", kontroller),
        ("KARAR", karar_motoru),
        ("PANO", pano),
        ("SENARYO_DUYARLILIK", senaryo_duyarlilik),
        ("RAPOR", rapor),
        ("ORNEK_VERI", ornek_veri),
        ("DEGISIKLIK_KAYDI", degisiklik_kaydi),
        ("LISTELER", listeler),
        ("AYARLAR", ayarlar),
        ("KILAVUZ", kilavuz),
    ]
    for ad, fn in sayfalar:
        ws2 = wb.create_sheet()
        fn(ws2)

    # Ad tanımları — motor çıktıları
    t = 6 + len(KIYMET_TURLERI)
    ad_ekle(wb, "YenidegerNetDefterToplam", f"HESAP!$B${t + 1}")
    ad_ekle(wb, "YenidegerNetFaydaToplam", f"HESAP!$B${t + 2}")
    ad_ekle(wb, "YenidegerFonToplam", f"HESAP!$B${t + 3}")
    ad_ekle(wb, "YenidegerDegerArtisiToplam", f"HESAP!$B${t + 4}")
    ad_ekle(wb, "YenidegerYillikTasarruf", f"HESAP!$B${t + 5}")
    ad_ekle(wb, "YenidegerTasarrufBdToplam", f"HESAP!$B${t + 6}")
    ad_ekle(wb, "YenidegerKalanOmurOrtalama", f"HESAP!$B${t + 7}")
    ad_ekle(wb, "YenidegerKiyametAdedi", f"HESAP!$B${t + 8}")
    ad_ekle(wb, "YenidegerNetFaydaVerimlilik", f"HESAP!$B${t + 9}")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "YenidegerFonOran": 11, "YenidegerKurumlarVergiOran": 12, "YenidegerIskontoOrani": 13,
        "YenidegerIyiIskonto": 14, "YenidegerBazIskonto": 15, "YenidegerKotuIskonto": 16,
        "YenidegerKritikIskonto": 17, "YenidegerNbdEsik": 18, "YenidegerArtisEsik": 19,
        "YenidegerFonEsik": 20, "YenidegerTornadoIskonto": 21, "YenidegerTornadoFon": 22,
        "YenidegerTahminCarpan": 23, "YenidegerKaliteIyi": 24, "YenidegerKaliteOrta": 25,
        "YenidegerAksiyon1_ACIKLAMA": 26, "YenidegerAksiyon2_ACIKLAMA": 27,
        "YenidegerAksiyon3_ACIKLAMA": 28, "YenidegerP90Oran": 29, "YenidegerTemelKolonSayisi": 30,
    }
    for i in range(1, 13):
        ayar_satirlari[f"YenidegerPanelSatir{i}"] = 31 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "YenidegerToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "YenidegerDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1ArtisOrani": 6, "modulT1ArtisOraniYorum": 6,
        "modulT2OrtalamaKatsayi": 7, "modulT2OrtalamaKatsayiYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2EnYuksekFayda": 9, "modulO2EnYuksekFaydaYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1FaydaTahmin": 13, "modulI1FaydaTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeKiyametTurleri", "LISTELER!$A$7:$A$12")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "YenidenDegerlemeVergiTasarrufAnalizi.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
