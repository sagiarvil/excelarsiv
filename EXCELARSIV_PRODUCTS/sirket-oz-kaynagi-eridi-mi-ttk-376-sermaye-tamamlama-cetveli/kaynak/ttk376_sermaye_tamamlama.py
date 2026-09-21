"""
Şirket Öz Kaynağı Eridi Mi? (TTK 376 Sermaye Tamamlama Cetveli) — üretim betiği.
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
    KRITIK,
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

URUN_AD = "Şirket Öz Kaynağı Eridi Mi? (TTK 376 Sermaye Tamamlama Cetveli)"
SURUM = "1.0.0"
RENK = "0F2742"

BORC_TURLERI = ["Banka Kredisi", "Tedarikçi Borcu", "Finansal Kiralama", "Tahvil", "Ortak Borcu", "Diğer"]
EVET_HAYIR = ["Evet", "Hayır"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Aylık bilanço verilerinizi girin; öz kaynağın sermayeyi karşılama durumunu, "
                "TTK 376 kapsamındaki sermaye kaybı eşiklerini, borca batıklığı ve yapılması "
                "gerekenleri otomatik hesaplayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Aylık öz kaynak bileşenlerinden gerçek öz kaynağı ve sermaye kaybını otomatik hesaplar",
        "TTK 376'daki 1/3 ve 2/3 sermaye kaybı eşiklerini canlı kontrol eder (hâlâ sağlıklı mı?)",
        "Borca batıklık tespiti ve yönetim kurulu bildirim yükümlülüğü uyarısı",
        "Sermaye artırımı / azaltımı / infisah senaryolarında öz kaynak ve kayıp oranı karşılaştırması",
        "UYGUN / İNCELE / DURDUR karar kapısı ve yönetici raporu",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Öz kaynağı eriyen ve TTK 376 riski taşıyan anonim / limited şirketler",
        "Genel kurul toplantısı öncesi sermaye tamamlama veya azaltım kararı alacak yönetim kurulları",
        "Borca batıklık bildirimi ve mali tablo sunumu hazırlayan mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. DONEM_BILANCO sayfasına aylık bilanço bileşenlerinizi girin. "
                 "3. BORC_LISTESI sayfasına borçlarınızı ekleyin. "
                 "4. PANO ve RAPOR sayfalarından öz kaynak durumunu ve kararı izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Aylık bilanço verilerini girin",
         "DONEM_BILANCO sayfasındaki sarı hücrelere her ayın ödenmiş sermaye, yedek, geçmiş yıl "
         "kâr/zarar, dönem kâr/zarar, borç ve aktif değerlerini girin. Öz kaynak otomatik hesaplanır."),
        ("Adım 2 — Borç listesini girin",
         "BORC_LISTESI sayfasındaki sarı hücrelere borçlarınızı ve türlerini girin; yoğunlaşma "
         "analizi bu listeden beslenir."),
        ("Adım 3 — Sermaye kaybını izleyin",
         "HESAP sayfası öz kaynağı, kayıp tutarını, kayıp oranını ve 1/3 ile 2/3 eşiklerini üretir."),
        ("Adım 4 — Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye ve genel kurula sunulabilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    hatalar = [
        "Ödenmiş sermayeyi cari değer yerine tarihsel değerle girmek (kayıp oranı bozulur)",
        "Dönem zararını negatif işaretle girmek (önce ayrı bir zarar kolonu kullanılır)",
        "Borç/aktif tutarlarını dönem sonu bilançosuyla mutabakat ettirmemek",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; hukuki ve mali müşavir görüşü yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def donem_bilanco(ws):
    sayfa_hazirla(ws, "DONEM_BILANCO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem Bilanço Bileşenleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere her ayın bilanço bileşenlerini girin. Öz kaynak, kayıp tutarı ve "
                "kayıp oranı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ay", "Ödenmiş Sermaye (₺)", "Yedekler (₺)", "Geçmiş Yıl Kârı (₺)",
                "Geçmiş Yıl Zararı (₺)", "Dönem Kârı (₺)", "Dönem Zararı (₺)",
                "Toplam Borç (₺)", "Toplam Aktif (₺)", "Öz Kaynak (₺)",
                "Kayıp Tutarı (₺)", "Kayıp Oranı", "TTK 376 Durum"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Öz Kaynak (₺)": '=IF(tblBilanco[[#This Row],[Ay]]="","",tblBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]+tblBilanco[[#This Row],[Yedekler (₺)]]+tblBilanco[[#This Row],[Geçmiş Yıl Kârı (₺)]]-tblBilanco[[#This Row],[Geçmiş Yıl Zararı (₺)]]+tblBilanco[[#This Row],[Dönem Kârı (₺)]]-tblBilanco[[#This Row],[Dönem Zararı (₺)]])',
        "Kayıp Tutarı (₺)": '=IF(tblBilanco[[#This Row],[Ay]]="","",MAX(0,tblBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]-tblBilanco[[#This Row],[Öz Kaynak (₺)]]))',
        "Kayıp Oranı": '=IF(tblBilanco[[#This Row],[Ay]]="","",IF(tblBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]=0,0,tblBilanco[[#This Row],[Kayıp Tutarı (₺)]]/tblBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]))',
        "TTK 376 Durum": '=IF(tblBilanco[[#This Row],[Ay]]="","",IF(tblBilanco[[#This Row],[Öz Kaynak (₺)]]<0,"BORCA BATIK",IF(tblBilanco[[#This Row],[Toplam Aktif (₺)]]<tblBilanco[[#This Row],[Toplam Borç (₺)]],"BORCA BATIK",IF(tblBilanco[[#This Row],[Kayıp Oranı]]>=TtkUcteIkiEsik,"2/3 KAYIP (KRİTİK)",IF(tblBilanco[[#This Row],[Kayıp Oranı]]>=TtkUcteBirEsik,"1/3 KAYIP (DİKKAT)","UYGUN")))))',
    }
    tablo_ekle(ws, "tblBilanco", "A5:M1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeAylar", "A6:A1004",
              baslik="Ay", mesaj="Listeden dönem ayını seçin (ör. 2026-01).",
              hata_baslik="Geçersiz ay", hata_mesaj="Listede olmayan ay giremezsiniz.")
    for harf, ad, ust in (("B", "Ödenmiş Sermaye", "1000000000000"), ("C", "Yedekler", "1000000000000"),
                          ("D", "Geçmiş Yıl Kârı", "1000000000000"), ("E", "Geçmiş Yıl Zararı", "1000000000000"),
                          ("F", "Dönem Kârı", "1000000000000"), ("G", "Dönem Zararı", "1000000000000"),
                          ("H", "Toplam Borç", "1000000000000"), ("I", "Toplam Aktif", "1000000000000")):
        dogrulama(ws, "decimal", "0", f"{harf}6:{harf}1004",
                  baslik=ad, mesaj=f"Bu dönem sonu {ad} tutarını TL olarak girin.",
                  hata_baslik="Geçersiz değer", hata_mesaj=f"0 ile {ust} arasında olmalı.",
                  isaret="between", f2=ust)
    for satir in range(6, 1004):
        for kolon in range(2, 13):
            ws.cell(row=satir, column=kolon).number_format = TL if kolon in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11) else YÜZDE
    ornek = [
        ("2026-01", 5000000, 800000, 200000, 0, 350000, 0, 6200000, 9800000),
        ("2026-02", 5000000, 800000, 200000, 0, 180000, 0, 6500000, 9600000),
        ("2026-03", 5000000, 800000, 200000, 0, 90000, 0, 6900000, 9400000),
        ("2026-04", 5000000, 800000, 200000, 0, 0, 150000, 7200000, 9100000),
        ("2026-05", 5000000, 800000, 200000, 0, 0, 320000, 7500000, 8800000),
        ("2026-06", 5000000, 800000, 200000, 0, 0, 480000, 7900000, 8500000),
        ("2026-07", 5000000, 800000, 200000, 0, 0, 610000, 8300000, 8200000),
        ("2026-08", 5000000, 800000, 200000, 0, 0, 750000, 8700000, 7900000),
        ("2026-09", 5000000, 800000, 200000, 0, 0, 860000, 9100000, 7600000),
        ("2026-10", 5000000, 800000, 200000, 0, 0, 980000, 9500000, 7300000),
        ("2026-11", 5000000, 800000, 200000, 0, 0, 1120000, 9900000, 7000000),
        ("2026-12", 5000000, 800000, 200000, 0, 0, 1250000, 10300000, 6700000),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            ws.cell(row=i, column=k).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    sabitle(ws, "A6")
    genislik(ws, {"A": 34, "B": 18, "C": 14, "D": 18, "E": 18, "F": 14, "G": 14,
                  "H": 16, "I": 16, "J": 16, "K": 14, "L": 12, "M": 20})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def borc_listesi(ws):
    sayfa_hazirla(ws, "BORC_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Borç Listesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere dönem borçlarınızı girin. Yoğunlaşma (HHI) ve anomali analizleri "
                "bu listeden beslenir; borca batıklık kontrolünde dönem sonu toplam borç kullanılır.",
      yazi=GRİ, kaydir=True)
    sutunlar = ["Borç Kodu", "Borç Türü", "Alacaklı", "Kullanım Tarihi", "Vade Tarihi",
                "Borç Tutarı (₺)", "Yıllık Faiz Oranı", "Teminat Var mı?", "Açıklama",
                "Faiz Yükü (₺)"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Faiz Yükü (₺)": '=IF(tblBorc[[#This Row],[Borç Kodu]]="","",tblBorc[[#This Row],[Borç Tutarı (₺)]]*tblBorc[[#This Row],[Yıllık Faiz Oranı]])',
    }
    tablo_ekle(ws, "tblBorc", "A5:J1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeBorcTurleri", "B6:B1004",
              baslik="Borç Türü", mesaj="Listeden borç türünü seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "list", "ListeEvetHayir", "H6:H1004",
              baslik="Teminat Var mı?", mesaj="Evet veya Hayır seçin.",
              hata_baslik="Geçersiz seçim", hata_mesaj="Evet veya Hayır girin.")
    dogrulama(ws, "date", "01.01.2024", "D6:D1004",
              baslik="Kullanım Tarihi", mesaj="GG.AA.YYYY biçiminde borç kullanım tarihi girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "date", "01.01.2024", "E6:E1004",
              baslik="Vade Tarihi", mesaj="GG.AA.YYYY biçiminde vade tarihi girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2030 aralığında olmalı.",
              isaret="between", f2="31.12.2030")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Borç Tutarı", mesaj="Borç tutarını TL olarak girin.",
              hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "G6:G1004",
              baslik="Faiz Oranı", mesaj="Yıllık faiz oranını ondalık girin (ör. %25 için 0,25).",
              hata_baslik="Geçersiz oran", hata_mesaj="0 ile 2 arasında olmalı.",
              isaret="between", f2="2")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=4).number_format = TARİH
        ws.cell(row=satir, column=5).number_format = TARİH
        ws.cell(row=satir, column=6).number_format = TL
        ws.cell(row=satir, column=7).number_format = YÜZDE
        ws.cell(row=satir, column=10).number_format = TL
    ornek = [
        ("B-1001", "Banka Kredisi", "ABC Bank", "2026-01-15", "2027-01-15", 3500000, 0.30, "Evet", "İşletme kredisi"),
        ("B-1002", "Tedarikçi Borcu", "Tedarikçi X", "2026-02-01", "2026-08-01", 1200000, 0.00, "Hayır", "Vadeli alım"),
        ("B-1003", "Tahvil", "Kurumsal Tahvil", "2026-03-10", "2028-03-10", 2000000, 0.28, "Hayır", "Kurumsal tahvil"),
        ("B-1004", "Ortak Borcu", "Ortak A", "2026-04-05", "2027-04-05", 1800000, 0.20, "Hayır", "Ortak avansı"),
        ("B-1005", "Finansal Kiralama", "Kiralama A.Ş.", "2026-05-20", "2028-05-20", 900000, 0.24, "Evet", "Araç kiralaması"),
        ("B-1006", "Banka Kredisi", "XYZ Bank", "2026-06-01", "2027-06-01", 1500000, 0.32, "Evet", "Yatırım kredisi"),
    ]
    for i, satir in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = satir[0]
        ws.cell(row=i, column=2).value = satir[1]
        ws.cell(row=i, column=3).value = satir[2]
        ws.cell(row=i, column=4).value = date(*map(int, satir[3].split("-")))
        ws.cell(row=i, column=5).value = date(*map(int, satir[4].split("-")))
        for k in (6, 7, 8, 9):
            ws.cell(row=i, column=k).value = satir[k - 1]
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 20, "C": 22, "D": 14, "E": 14, "F": 18, "G": 14, "H": 14, "I": 22, "J": 16})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "TTK 376 Sermaye Tamamlama Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet, dönem bilanço ve borç tablolarından otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    sutunlar = ["Borç Türü", "Borç Adedi", "Toplam Borç (₺)", "Ortalama Faiz Oranı",
                "Teminatlı Borç (₺)", "Borç Payı"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(BORC_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=COUNTIFS(tblBorc[Borç Türü],$A{r})', sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblBorc[Borç Tutarı (₺)],tblBorc[Borç Türü],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=IF($B${r}=0,0,AVERAGEIFS(tblBorc[Yıllık Faiz Oranı],tblBorc[Borç Türü],$A{r}))',
          sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblBorc[Borç Tutarı (₺)],tblBorc[Borç Türü],$A{r},tblBorc[Teminat Var mı?],"Evet")',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=IF($C${11}=0,0,C{r}/$C${11})', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        for k in range(2, 7):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: {tur} borç metriği | "
                       f"Neden önemli: Borç yapısını ve yoğunlaşmayı gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    t = 6 + len(BORC_TURLERI)
    h(ws, t, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    for k in range(2, 6):
        h(ws, t, k, f"=SUM({chr(64+k)}6:{chr(64+k)}{t-1})", sayi=(TL if k in (3, 5) else CATI),
          yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t, 6, "=SUM(F6:F11)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 1, 1, "Dönem Sonu Ödenmiş Sermaye", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 1, 2, "=IFERROR(INDEX(tblBilanco[Ödenmiş Sermaye (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 2, 1, "Dönem Sonu Yedekler", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 2, 2, "=IFERROR(INDEX(tblBilanco[Yedekler (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 3, 1, "Dönem Sonu Geçmiş Yıl Kârı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 3, 2, "=IFERROR(INDEX(tblBilanco[Geçmiş Yıl Kârı (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 4, 1, "Dönem Sonu Geçmiş Yıl Zararı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 4, 2, "=IFERROR(INDEX(tblBilanco[Geçmiş Yıl Zararı (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 5, 1, "Dönem Sonu Dönem Kârı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 5, 2, "=IFERROR(INDEX(tblBilanco[Dönem Kârı (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 6, 1, "Dönem Sonu Dönem Zararı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 6, 2, "=IFERROR(INDEX(tblBilanco[Dönem Zararı (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 7, 1, "Dönem Sonu Öz Kaynak Toplamı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 7, 2, f"=B{t + 1}+B{t + 2}+B{t + 3}-B{t + 4}+B{t + 5}-B{t + 6}",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 8, 1, "Sermaye Kaybı Tutarı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 8, 2, f"=MAX(0,B{t + 1}-B{t + 7})", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 9, 1, "Kayıp Oranı (Kayıp / Ödenmiş Sermaye)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 9, 2, f"=IF(B{t + 1}=0,0,B{t + 8}/B{t + 1})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 10, 1, "1/3 Sermaye Kaybı Eşiği (TTK 376)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 10, 2, f"=B{t + 1}*TtkUcteBirEsik", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 11, 1, "2/3 Sermaye Kaybı Eşiği (TTK 376)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 11, 2, f"=B{t + 1}*TtkUcteIkiEsik", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 12, 1, "Dönem Sonu Toplam Borç", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 12, 2, "=IFERROR(INDEX(tblBilanco[Toplam Borç (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 13, 1, "Dönem Sonu Toplam Aktif", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 13, 2, "=IFERROR(INDEX(tblBilanco[Toplam Aktif (₺)],COUNTA(tblBilanco[Ay])),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 14, 1, "Borca Batıklık Durumu", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 14, 2, f"=IF(OR(B{t + 7}<0,B{t + 13}<B{t + 12}),\"BORCA BATIK\",\"BORCA BATIK DEĞİL\")",
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 15, 1, "Öz Kaynak / Aktif Oranı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 15, 2, f"=IF(B{t + 13}=0,0,B{t + 7}/B{t + 13})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 16, 1, "İlk-Son Ay Öz Kaynak Farkı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 16, 2, '=IFERROR(INDEX(tblBilanco[Öz Kaynak (₺)],COUNTA(tblBilanco[Ay]))-INDEX(tblBilanco[Öz Kaynak (₺)],1),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 17, 1, "Uç Değer Sinyali", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 17, 2, f"=IF(OR(B{t + 12}>TtkAzamiBilancoTutari,B{t + 13}>TtkAzamiBilancoTutari),\"AŞIM VAR\",\"NORMAL\")",
      yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, t + 19, "Hesap satırları yalnızca dönem bilanço ve borç tablolarından beslenir; korumalıdır.")
    genislik(ws, {"A": 42, "B": 22, "C": 20, "D": 20, "E": 20, "F": 12})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Öz Kaynak Trendi (Son/İlk Ay)",
         "=IFERROR(IF(INDEX(tblBilanco[Öz Kaynak (₺)],1)=0,0,INDEX(tblBilanco[Öz Kaynak (₺)],COUNTA(tblBilanco[Ay]))/INDEX(tblBilanco[Öz Kaynak (₺)],1)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Öz kaynak trendi ",TEXT(C6,"0.00"),"x — 1 altındaysa öz kaynak eriyor, eşikler yaklaşıyor")'),
        ("T4", "İlk-Son Ay Öz Kaynak Farkı",
         "=HESAP!B28",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İlk ile son ay öz kaynak farkı ",TEXT(C7,"₺ #,##0")," TL — negatif fark erime hızını gösterir")'),
        ("O1", "Aylık Kayıp Oranı Sapması (STDEV)",
         "=IFERROR(STDEV.P(tblBilanco[Kayıp Oranı]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Aylık kayıp oranlarının standart sapması ",TEXT(C8,"0.00%")," — sapma yüksekse erime hızı düzensizdir")'),
        ("O2", "Tornado: En Büyük Etki",
         "=MAX(SENARYO_DUYARLILIK!B12:SENARYO_DUYARLILIK!B13)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kayıp oranını en çok etkileyen değişkenin etkisi ",TEXT(C9,"0.00%")," — öncelik bu değişkenin yönetiminde")'),
        ("O3", "Senaryo Bant Genişliği (Kayıp Oranı)",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C9",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kötümser senaryo kayıp oranı farkı ",TEXT(C10,"0.00%")," — belirsizlik aralığını gösterir")'),
        ("O6", "Borç Türü Yoğunlaşması (HHI)",
         "=IFERROR(SUMPRODUCT((HESAP!$C$6:$C$11/SUM(HESAP!$C$6:$C$11))^2),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Borç yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça borç tek türde toplanır, kırılganlık artar")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=TtkKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Öz Kaynak Tahmini",
         "=PANO!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem öz kaynak tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Borç Tutarı Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblBorc[Borç Tutarı (₺)],TtkP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Borç tutarlarının %90 ı ",TEXT(C14,"₺ #,##0")," TL altında gerçekleşir — üst sınır finansman planı için kullanılabilir")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        satir = 6 + i
        h(ws, satir, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, satir, 2, ad, kaydir=True)
        h(ws, satir, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[satir].height = 30
    genislik(ws, {"A": 6, "B": 32, "C": 52, "D": 72})
    ws.sheet_view.showGridLines = True


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Giriş Kalite Kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki kontroller canlı formüllerle çalışır; sorun varsa uyarı üretir.",
      yazi=GRİ, kaydir=True)
    kontroller = [
        ("Boş ay satırı var mı?", "=IF(COUNTBLANK(tblBilanco[Ay])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Boş ödenmiş sermaye var mı?", "=IF(COUNTBLANK(tblBilanco[Ödenmiş Sermaye (₺)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Negatif değer var mı?", "=IF(COUNTIF(tblBilanco[Ödenmiş Sermaye (₺)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Ödenmiş sermaye sıfır mı?", "=IF(COUNTIF(tblBilanco[Ödenmiş Sermaye (₺)],\"=0\")>0,\"SIFIR VAR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Borç türü boş mu?", "=IF(COUNTIF(tblBorc[Borç Türü],\"\")>0,\"BOŞ VAR\",\"TEMİZ\")", NORMAL, KRITIK),
        ("Kayıp oranı 1/3 eşiğini aşıyor mu?", "=IF(HESAP!B21>TtkUcteBirEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Kayıp oranı 2/3 eşiğini aşıyor mu?", "=IF(HESAP!B21>TtkUcteIkiEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Borca batıklık var mı?", "=IF(HESAP!B26=\"BORCA BATIK\",\"VAR\",\"YOK\")", KRITIK, NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Boş ve negatif girişlerin oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblBilanco[Ay])*TtkDonemGirisKolon+COUNTA(tblBorc[Borç Kodu])*TtkBorcGirisKolon",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblBilanco[Ay])+COUNTA(tblBilanco[Ödenmiş Sermaye (₺)])+COUNTA(tblBilanco[Yedekler (₺)])+"
                 "COUNTA(tblBilanco[Geçmiş Yıl Kârı (₺)])+COUNTA(tblBilanco[Geçmiş Yıl Zararı (₺)])+"
                 "COUNTA(tblBilanco[Dönem Kârı (₺)])+COUNTA(tblBilanco[Dönem Zararı (₺)])+"
                 "COUNTA(tblBilanco[Toplam Borç (₺)])+COUNTA(tblBilanco[Toplam Aktif (₺)])+"
                 "COUNTA(tblBorc[Borç Kodu])+COUNTA(tblBorc[Borç Türü])+COUNTA(tblBorc[Borç Tutarı (₺)])+"
                 "COUNTA(tblBorc[Yıllık Faiz Oranı])+COUNTA(tblBorc[Teminat Var mı?])+COUNTA(tblBorc[Açıklama])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(TtkToplamGiris=0,0,TtkDoluGiris/TtkToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=TtkKaliteIyi,\"YÜKSEK\",IF(B19>=TtkKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 55})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — TTK 376 Sermaye Tamamlama", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sermaye kaybı eşikleri ve borca batıklık durumuna göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Dönem Sonu Öz Kaynak", yazi="333333")
    h(ws, 6, 2, "=HESAP!B19", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Sermaye Kaybı Tutarı", yazi="333333")
    h(ws, 7, 2, "=HESAP!B20", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Kayıp Oranı", yazi="333333")
    h(ws, 8, 2, "=HESAP!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Borca Batıklık Durumu", yazi="333333")
    h(ws, 9, 2, "=HESAP!B26", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblBilanco[Ay])=0,"VERİ YOK",IF(HESAP!B29="AŞIM VAR","İNCELE",IF(HESAP!B26="BORCA BATIK","DURDUR",IF(HESAP!B21>=TtkUcteIkiEsik,"DURDUR",IF(HESAP!B21>=TtkUcteBirEsik,"İNCELE","UYGUN")))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTA(tblBilanco[Ay])=0,"Veri girişi yapılmamış; karar üretilemez.",'\
                 'IF(HESAP!B29="AŞIM VAR","Bilanço tutarları gerçekçilik üst sınırının üzerinde; verilerin doğru girildiği doğrulanmalı.",'\
                 'IF(HESAP!B26="BORCA BATIK","Öz kaynak negatif veya aktifler borçları karşılamıyor; TTK 376 kapsamında yönetim kurulu mahkemeye bildirim yükümlülüğü doğar.",'\
                 'IF(HESAP!B21>=TtkUcteIkiEsik,"Sermayenin 2/3 ü kaybolmuş; genel kurul derhal toplanmalı, sermaye tamamlama veya infisah kararı alınmalı.",'\
                 'IF(HESAP!B21>=TtkUcteBirEsik,"Sermayenin 1/3 ü kaybolmuş; genel kurul toplanıp tamamlama veya azaltım kararı alınmalı.",'\
                 '"Öz kaynak sermayenin 1/3 ünden fazla erimemiş; TTK 376 kapsamında zorunlu önlem gerekmiyor.")))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=TtkAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki eşiklerden türetilir; "
                     "kesin hukuki sonuç değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Dönem Sonu Öz Kaynak", "=HESAP!B19", TL),
        (3, "Sermaye Kaybı", "=HESAP!B20", TL),
        (4, "Kayıp Oranı", "=HESAP!B21", YÜZDE),
        (5, "Öz Kaynak / Aktif", "=HESAP!B27", YÜZDE),
        (6, "Borca Batıklık", "=HESAP!B26", None),
        (7, "Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        if sayi:
            h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
        else:
            h(ws, 4, kolon, form, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Toplam Borç", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B24", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — borç türü bazlı
    h(ws, 10, 1, "Borç Türü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "Borç Adedi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 3, "Toplam Borç (₺)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 10, 6, "Teminatlı Borç (₺)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    for i, tur in enumerate(BORC_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i-5}", sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!C{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 6, f"=HESAP!E{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 4, "Aylık Öz Kaynak", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    for i in range(12):
        h(ws, 11 + i, 4, f"AY{i+1}", yazi=GRİ)
        h(ws, 11 + i, 5, f"=IFERROR(INDEX(tblBilanco[Öz Kaynak (₺)],AYARLAR!$B${13 + i}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Borç Türü Adet ve Tutar"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=16), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F6")
    g2 = PieChart()
    g2.title = "Borç Adedi Dağılımı"
    g2.add_data(Reference(ws, min_col=2, min_row=10, max_row=16), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U6")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Borç Türü Teminatlı Borç"
    g3.add_data(Reference(ws, min_col=6, min_row=10, max_row=16), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g3.height, g3.width = 8, 13
    ws.add_chart(g3, "F19")
    g4 = LineChart()
    g4.title = "Aylık Öz Kaynak Trendi"
    g4.add_data(Reference(ws, min_col=5, min_row=10, max_row=23), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=11, max_row=22))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U19")
    # Zaman serisi projeksiyonu — öz kaynak
    h(ws, 24, 1, "Gelecek 1 Dönem Öz Kaynak Tahmini (Tahmin Aralıklı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, '=ROUND(IFERROR(FORECAST.LINEAR(TtkPanelSatir12,tblBilanco[Öz Kaynak (₺)],ROW(tblBilanco[Öz Kaynak (₺)])-ROW(tblBilanco[[#Headers],[Öz Kaynak (₺)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Ortalama Aylık Öz Kaynak", yazi="333333")
    h(ws, 25, 2, "=ROUND(IFERROR(AVERAGE(tblBilanco[Öz Kaynak (₺)]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Öz Kaynak Standart Sapma", yazi="333333")
    h(ws, 26, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblBilanco[Öz Kaynak (₺)],TtkPanelSatir1):INDEX(tblBilanco[Öz Kaynak (₺)],TtkPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 27, 2, "=B25-B26*TtkTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 28, 2, "=B25+B26*TtkTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=24, max_row=28), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=25, max_row=28))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F33")
    g6 = LineChart()
    g6.title = "Borç Türü Toplam Borç"
    g6.add_data(Reference(ws, min_col=3, min_row=10, max_row=16), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=11, max_row=16))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U33")
    # Senaryo karşılaştırma verisi
    h(ws, 29, 3, "Senaryo Kayıp Oranı", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser"], 30):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i-23}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo Kayıp Oranı Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=29, max_row=32), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=30, max_row=32))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F39")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Öz Kaynak ve Sermaye Yapısı"
    g8.add_data(Reference(ws, min_col=2, min_row=3, max_row=4), titles_from_data=True)
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U39")
    baski_hazirla(ws, "A1:M56", URUN_AD)
    alt_bant(ws, 56, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 45, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İyimser, baz ve kötümser senaryolarda kayıp oranı; ardından hangi değişkenin "
                "sonucu en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    senaryolar = [
        ("İYİMSER", "TtkIyiSenaryo"),
        ("BAZ", "TtkBazSenaryo"),
        ("KÖTÜMSER", "TtkKotuSenaryo"),
    ]
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Dönem Zararı Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Kayıp Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "Dönem Sonu Öz Kaynak", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, carpan) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=IF(SUM(tblBilanco[Dönem Zararı (₺)])=0,0,{carpan})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=IF(HESAP!B13=0,0,MAX(0,HESAP!B13-(HESAP!B19-HESAP!B18*({carpan}-1)))/HESAP!B13)",
          sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!B19-HESAP!B18*({carpan}-1)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Tornado: Kayıp Oranına Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tornadolar = [
        ("Ödenmiş Sermaye +%1", "=IF(HESAP!B13=0,0,ABS(HESAP!B21-MAX(0,HESAP!B13*(1+TtkTornadoZarar)-HESAP!B19)/(HESAP!B13*(1+TtkTornadoZarar))))", "Ödenmiş Sermaye"),
        ("Dönem Zararı +%1", "=IF(HESAP!B13=0,0,ABS(HESAP!B21-MAX(0,HESAP!B13-(HESAP!B19-HESAP!B18*TtkTornadoZarar))/HESAP!B13))", "Dönem Zararı"),
    ]
    for i, (ad, form, etiket) in enumerate(tornadolar, 12):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IF(B12>=B13,\"Ödenmiş Sermaye\",\"Dönem Zararı\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 17, "Senaryo çarpanları AYARLAR'daki parametrelerden türetilir.")
    genislik(ws, {"A": 45, "B": 30, "C": 20, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "ŞİRKET ÖZ KAYNAĞI — TTK 376 SERMAYE TAMAMLAMA YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Dönem Sonu Öz Kaynak", "=HESAP!B19", TL),
        ("Sermaye Kaybı Tutarı", "=HESAP!B20", TL),
        ("Kayıp Oranı", "=HESAP!B21", YÜZDE),
        ("1/3 Sermaye Kaybı Eşiği", "=HESAP!B22", TL),
        ("2/3 Sermaye Kaybı Eşiği", "=HESAP!B23", TL),
        ("Borca Batıklık Durumu", "=HESAP!B26", None),
        ("Toplam Borç", "=HESAP!B24", TL),
        ("Öz Kaynak / Aktif Oranı", "=HESAP!B27", YÜZDE),
        ("Borç Yoğunlaşması (HHI)", "=ANALITIK_MOTOR!C11", "0.000"),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        if sayi:
            h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
        else:
            h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "Öz kaynak, ödenmiş sermaye + yedekler + geçmiş yıl kârı - geçmiş yıl zararı + dönem kârı - dönem zararı olarak hesaplanır.",
        "Sermaye kaybı eşikleri (1/3 ve 2/3) TTK 376 kapsamında değerlendirilir; kesin hukuki sonuç, genel kurul kararı ve tescil sürecine bağlıdır.",
        "Borca batıklık: öz kaynak negatif ise veya aktifler borçları karşılamıyorsa işaretlenir; bu durumda yönetim kurulu bildirim yükümlülüğü değerlendirilmelidir.",
        "Bu dosya hukuki ve mali müşavir görüşü yerine geçmez; mahkeme ve genel kurul süreçleri için uzman desteği gerektirir.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 18):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=TtkAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
        ws.merge_cells(start_row=25 + i, start_column=1, end_row=25 + i, end_column=12)
    h(ws, 29, 1, "Bu rapor karar destek amaçlıdır; hukuki tavsiye niteliği taşımaz.",
      yazi=GRİ, boyut=9, italik=True, kaydir=True)
    baski_hazirla(ws, "A1:L29", URUN_AD)
    genislik(ws, {"A": 45, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek Veri (Demo)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Gerçek girişlere başlamadan önce demo veri ile ürünü deneyin. "
                "Bu satırlar örnektir; gerçek verilerle değiştirin.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ay", "Ödenmiş Sermaye (₺)", "Yedekler (₺)", "Geçmiş Yıl Kârı (₺)",
                "Geçmiş Yıl Zararı (₺)", "Dönem Kârı (₺)", "Dönem Zararı (₺)",
                "Toplam Borç (₺)", "Toplam Aktif (₺)", "Öz Kaynak (₺)",
                "Kayıp Tutarı (₺)", "Kayıp Oranı", "TTK 376 Durum"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Öz Kaynak (₺)": '=IF(tblOrnekBilanco[[#This Row],[Ay]]="","",tblOrnekBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]+tblOrnekBilanco[[#This Row],[Yedekler (₺)]]+tblOrnekBilanco[[#This Row],[Geçmiş Yıl Kârı (₺)]]-tblOrnekBilanco[[#This Row],[Geçmiş Yıl Zararı (₺)]]+tblOrnekBilanco[[#This Row],[Dönem Kârı (₺)]]-tblOrnekBilanco[[#This Row],[Dönem Zararı (₺)]])',
        "Kayıp Tutarı (₺)": '=IF(tblOrnekBilanco[[#This Row],[Ay]]="","",MAX(0,tblOrnekBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]-tblOrnekBilanco[[#This Row],[Öz Kaynak (₺)]]))',
        "Kayıp Oranı": '=IF(tblOrnekBilanco[[#This Row],[Ay]]="","",IF(tblOrnekBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]=0,0,tblOrnekBilanco[[#This Row],[Kayıp Tutarı (₺)]]/tblOrnekBilanco[[#This Row],[Ödenmiş Sermaye (₺)]]))',
        "TTK 376 Durum": '=IF(tblOrnekBilanco[[#This Row],[Ay]]="","",IF(tblOrnekBilanco[[#This Row],[Öz Kaynak (₺)]]<0,"BORCA BATIK",IF(tblOrnekBilanco[[#This Row],[Toplam Aktif (₺)]]<tblOrnekBilanco[[#This Row],[Toplam Borç (₺)]],"BORCA BATIK",IF(tblOrnekBilanco[[#This Row],[Kayıp Oranı]]>=TtkUcteIkiEsik,"2/3 KAYIP (KRİTİK)",IF(tblOrnekBilanco[[#This Row],[Kayıp Oranı]]>=TtkUcteBirEsik,"1/3 KAYIP (DİKKAT)","UYGUN")))))',
    }
    tablo_ekle(ws, "tblOrnekBilanco", "A5:M1004", sutunlar, formuller)
    ornek = [
        ("2026-01", 5000000, 800000, 200000, 0, 350000, 0, 6200000, 9800000),
        ("2026-02", 5000000, 800000, 200000, 0, 180000, 0, 6500000, 9600000),
        ("2026-03", 5000000, 800000, 200000, 0, 90000, 0, 6900000, 9400000),
        ("2026-04", 5000000, 800000, 200000, 0, 0, 150000, 7200000, 9100000),
        ("2026-05", 5000000, 800000, 200000, 0, 0, 320000, 7500000, 8800000),
        ("2026-06", 5000000, 800000, 200000, 0, 0, 480000, 7900000, 8500000),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            h(ws, i, k, deger)
        for k in range(2, 12):
            ws.cell(row=i, column=k).number_format = TL
        ws.cell(row=i, column=12).number_format = YÜZDE
    h(ws, 3, 14, "Borç Listesi (Demo)", kalin=True, boyut=12, yazi=KOYU_LACIVERT, kaydir=True)
    sutunlar2 = ["Borç Kodu", "Borç Türü", "Alacaklı", "Borç Tutarı (₺)", "Yıllık Faiz Oranı",
                 "Teminat Var mı?", "Açıklama", "Faiz Yükü (₺)"]
    baslik_satiri(ws, 5, sutunlar2, basla=14)
    formuller2 = {
        "Faiz Yükü (₺)": '=IF(tblOrnekBorc[[#This Row],[Borç Kodu]]="","",tblOrnekBorc[[#This Row],[Borç Tutarı (₺)]]*tblOrnekBorc[[#This Row],[Yıllık Faiz Oranı]])',
    }
    tablo_ekle(ws, "tblOrnekBorc", "N5:U1004", sutunlar2, formuller2)
    ornek2 = [
        ("B-1001", "Banka Kredisi", "ABC Bank", 3500000, 0.30, "Evet", "İşletme kredisi"),
        ("B-1002", "Tedarikçi Borcu", "Tedarikçi X", 1200000, 0.00, "Hayır", "Vadeli alım"),
        ("B-1003", "Tahvil", "Kurumsal Tahvil", 2000000, 0.28, "Hayır", "Kurumsal tahvil"),
        ("B-1004", "Ortak Borcu", "Ortak A", 1800000, 0.20, "Hayır", "Ortak avansı"),
        ("B-1005", "Finansal Kiralama", "Kiralama A.Ş.", 900000, 0.24, "Evet", "Araç kiralaması"),
        ("B-1006", "Banka Kredisi", "XYZ Bank", 1500000, 0.32, "Evet", "Yatırım kredisi"),
    ]
    for i, satir in enumerate(ornek2, 6):
        for k, deger in enumerate(satir, 0):
            h(ws, i, 14 + k, deger)
        ws.cell(row=i, column=17).number_format = TL
        ws.cell(row=i, column=18).number_format = YÜZDE
        ws.cell(row=i, column=21).number_format = TL
    genislik(ws, {"A": 12, "B": 18, "C": 14, "D": 18, "E": 18, "F": 14, "G": 14,
                  "H": 16, "I": 16, "N": 16, "O": 14, "P": 12, "Q": 20,
                  "R": 24, "S": 20, "T": 18, "U": 16})
    alt_bant(ws, 1006, "Demo veriler rapor sonuçlarını örnekler; satın alma sonrası silinerek "
                       "gerçek veriler girilir.")


def degisiklik_kaydi(ws):
    sayfa_hazirla(ws, "DEGISIKLIK_KAYDI", "999999", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Değişiklik Kaydı", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 4, ["Tarih", "Sürüm", "Değişiklik", "Neden", "Yapan"])
    kayitlar = [
        ("01.08.2026", SURUM, "İlk üretim", "Ürünün yayına hazırlanması", "ExcelArşiv"),
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
    h(ws, 6, 3, "Ay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 7 + i, 3, f"2026-{i+1:02d}", zemin=GIRIS_SARI)
    h(ws, 6, 5, "Evet / Hayır", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, secim in enumerate(EVET_HAYIR, 7):
        h(ws, i, 5, secim, zemin=GIRIS_SARI)
    genislik(ws, {"A": 18, "C": 12, "E": 14})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Eşikleri ve çarpanları kendi işletmenize göre "
                "güncelleyin. Değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"], basla=1)
    YURURLUK = date(2026, 1, 1)
    DOGRULAMA = date(2026, 8, 10)
    ayarlar = [
        (7, "FirmaUnvan", "Örnek Firma A.Ş.", "Metin", "", "", "Rapor ve kapakta görünen firma adı.", "Kullanıcı"),
        (8, "RaporTarihi", date(2026, 8, 8), "Tarih", "", "", "Tüm hesapların ve kararın alındığı tarih.", "Kullanıcı"),
        (9, "RaporHazirlayan", "Finans Direktörü", "Metin", "", "", "Raporu hazırlayan kişi/rol.", "Kullanıcı"),
        (10, "DosyaSurumu", "1.0.0", "Metin", "", "", "Ürün sürüm bilgisi; rapor alt bilgisinde görünür.", "Kullanıcı"),
        (11, "TtkUcteBirEsik", 0.333333, "Oran", 0.05, 0.5, "TTK 376 1/3 sermaye kaybı eşiği; altına düşen öz kaynak işlemi tetikler.", "Mevzuat"),
        (12, "TtkUcteIkiEsik", 0.666667, "Oran", 0.3, 1, "TTK 376 2/3 sermaye kaybı eşiği; bu eşikte genel kurul derhal toplanır.", "Mevzuat"),
        (13, "TtkTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (14, "TtkIyiSenaryo", 0.6, "Oran", 0.1, 1, "İyimser senaryoda dönem zararı bu çarpanla hesaplanır (0.6 = %40 azalış).", "Varsayım"),
        (15, "TtkBazSenaryo", 1.0, "Oran", 0.1, 2, "Baz senaryoda dönem zararı bu çarpanla hesaplanır.", "Varsayım"),
        (16, "TtkKotuSenaryo", 1.5, "Oran", 0.1, 3, "Kötümser senaryoda dönem zararı bu çarpanla hesaplanır (1.5 = %50 artış).", "Varsayım"),
        (17, "TtkTornadoZarar", 0.01, "Oran", 0, 0.2, "Tornado analizinde dönem zararının %1 artış etkisi.", "Varsayım"),
        (18, "TtkKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (19, "TtkKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (20, "TtkAksiyon1_ACIKLAMA",
         "1/3 eşiğine yaklaşılıyorsa ödenmiş sermayeyi koruyacak şekilde ortaklardan sermaye taahhüdü veya borç/kâr transferi planlayın.", "Metin", "", "",
         "İNCELE kararında gösterilen ilk öneri.", "Varsayım"),
        (21, "TtkAksiyon2_ACIKLAMA",
         "2/3 kayıp durumunda genel kurulu derhal toplayın; sermaye artırımı, azaltım veya infisah kararı için hukuki ve mali danışman desteği alın.", "Metin", "", "",
         "DURDUR kararında gösterilen ikinci öneri.", "Varsayım"),
        (22, "TtkAksiyon3_ACIKLAMA",
         "Borca batıklıkta aktif değerlemesini yeniden yapın; bilanço güncellemesi ve mahkeme bildirimi sürecini mali müşavirle birlikte yürütün.", "Metin", "", "",
         "DURDUR kararında gösterilen üçüncü öneri.", "Varsayım"),
        (23, "TtkP90Oran", 0.9, "Oran", 0.5, 1, "Borç tutarı dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (24, "TtkAzamiBilancoTutari", 100000000000, "TL", 0, 100000000000000,
         "Gerçekçilik üst sınırı; toplam borç bu tutarı aşarsa uç değer işareti üretilir ve karar İNCELE kalır.", "Varsayım"),
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
        satir = 25 + i
        h(ws, satir, 1, f"TtkPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki aylık öz kaynak grafiği serisinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki aylık öz kaynak grafiğini besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    for satir, anahtar, deger, aciklama in [
        (38, "TtkDonemGirisKolon", 9, "DONEM_BILANCO tablosunda satır başına manuel giriş kolonu sayısı."),
        (39, "TtkBorcGirisKolon", 9, "BORC_LISTESI tablosunda satır başına manuel giriş kolonu sayısı."),
    ]:
        h(ws, satir, 1, anahtar, kalin=True, zemin=ACIK_GRI, boyut=9, kaydir=True)
        hb = h(ws, satir, 2, deger, yazi=GIRIS_YAZI, zemin=GIRIS_SARI, kalin=True)
        hb.number_format = CATI
        h(ws, satir, 3, "Adet", boyut=9, yazi=GRİ)
        h(ws, satir, 4, 1, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 5, 100, boyut=9, yazi=GRİ, sayi=CATI)
        h(ws, satir, 6, aciklama, boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        dv = DataValidation(type="decimal", operator="between", formula1="1",
                            formula2="100", allow_blank=False,
                            promptTitle=anahtar, prompt=aciklama,
                            showErrorMessage=True,
                            errorTitle="Değer aralık dışı", error=aciklama)
        ws.add_data_validation(dv)
        dv.add(f"B{satir}")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 30, "B": 22, "C": 12, "D": 10, "E": 10, "F": 60, "G": 14, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "DONEM_BILANCO sayfasındaki sarı hücrelere aylık bilanço bileşenlerinizi (ödenmiş sermaye, "
         "yedekler, geçmiş yıl kâr/zarar, dönem kâr/zarar, borç ve aktif) girin; BORC_LISTESI sayfasına "
         "borçlarınızı ve türlerini ekleyin."),
        ("Hesap katmanı",
         "HESAP sayfası öz kaynağı, sermaye kaybı tutarını, kayıp oranını ve TTK 376'daki 1/3 ile 2/3 "
         "eşiklerini otomatik özetler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR genel kurul ve yönetim kuruluna "
         "sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki 1/3 ve 2/3 eşiklerini ve senaryo çarpanlarını mali müşavirinize göre "
         "güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; hukuki ve mali müşavir görüşü, genel kurul kararı "
                 "ve tescil süreci yerine geçmez.", yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "Tahmin Çarpanı (standart sapma katı)",
        "İyimser, Baz ve Kötümser Senaryo Çarpanları",
        "Tornado Dönem Zararı Etkisi",
        "Kalite Skoru İyi ve Orta Eşikleri",
        "P90 Yüzdelik Oranı",
        "Aksiyon Önerisi Açıklamaları",
        "Panel Satır Numaraları",
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

    ws = wb["DONEM_BILANCO"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("J6:J1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("K6:K1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="between", formula=["0.333333", "0.666667"], font=f(amber)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="lessThan", formula=["0.333333"], font=f(yesil)))
    ws.conditional_formatting.add("M6:M1004", CellIsRule(operator="equal", formula=['"BORCA BATIK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("M6:M1004", CellIsRule(operator="equal", formula=['"2/3 KAYIP (KRİTİK)"'],
                                 font=f(amber), fill=d("FFF2CC")))
    ws.conditional_formatting.add("M6:M1004", CellIsRule(operator="equal", formula=['"1/3 KAYIP (DİKKAT)"'],
                                 font=f(amber), fill=d("FFF2CC")))
    ws.conditional_formatting.add("M6:M1004", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))

    ws = wb["BORC_LISTESI"]
    ws.conditional_formatting.add("F6:F1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("F6:F1004", CellIsRule(operator="greaterThan", formula=["5000000"], font=f(amber)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0.5"], font=f(amber)))
    ws.conditional_formatting.add("H6:H1004", CellIsRule(operator="equal", formula=['"Evet"'],
                                 font=f(yesil), fill=d("E2EFDA")))

    ws = wb["HESAP"]
    for harf in ("C", "D", "E"):
        ws.conditional_formatting.add(f"{harf}6:{harf}11",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B21:B21", CellIsRule(operator="greaterThan", formula=["0.333333"], font=f(amber)))
    ws.conditional_formatting.add("B21:B21", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))
    ws.conditional_formatting.add("B20:B20", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))
    ws.conditional_formatting.add("B26:B26", CellIsRule(operator="equal", formula=['"BORCA BATIK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "SIFIR VAR", "EKSİK VAR", "AŞIYOR", "YÜKSEK", "VAR"):
        ws.conditional_formatting.add("B6:B13", CellIsRule(operator="equal", formula=[f'"{deger}"'],
                                     font=f(kirmizi), fill=d("FDE9E9")))
    for deger in ("TEMİZ", "NORMAL", "YOK"):
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
    ws.conditional_formatting.add("B11:B11", CellIsRule(operator="equal", formula=['"VERİ YOK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B8:B8", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))

    ws = wb["PANO"]
    ws.conditional_formatting.add("B4:B4", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("C4:C4", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))
    ws.conditional_formatting.add("D4:D4", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))
    ws.conditional_formatting.add("D4:D4", CellIsRule(operator="between", formula=["0.333333", "0.666667"], font=f(amber)))
    ws.conditional_formatting.add("F4:F4", CellIsRule(operator="equal", formula=['"BORCA BATIK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E11:E22", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="between", formula=["0.333333", "0.666667"], font=f(amber)))
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="lessThan", formula=["0.333333"], font=f(yesil)))
    ws.conditional_formatting.add("D7:D9", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B12:B13", CellIsRule(operator="greaterThan", formula=["0.05"], font=f(amber)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("J6:J1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="greaterThan", formula=["0.666667"], font=f(kirmizi)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="between", formula=["0.333333", "0.666667"], font=f(amber)))
    ws.conditional_formatting.add("L6:L1004", CellIsRule(operator="lessThan", formula=["0.333333"], font=f(yesil)))
    ws.conditional_formatting.add("M6:M1004", CellIsRule(operator="equal", formula=['"BORCA BATIK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()
    kapak(ws)
    sayfalar = [
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("DONEM_BILANCO", donem_bilanco),
        ("BORC_LISTESI", borc_listesi),
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

    # Ad tanımları — HESAP motor çıktıları
    t = 6 + len(BORC_TURLERI)
    ad_ekle(wb, "TtkOdenmisSermaye", f"HESAP!$B${t + 1}")
    ad_ekle(wb, "TtkYedekler", f"HESAP!$B${t + 2}")
    ad_ekle(wb, "TtkGecmisYilKar", f"HESAP!$B${t + 3}")
    ad_ekle(wb, "TtkGecmisYilZarar", f"HESAP!$B${t + 4}")
    ad_ekle(wb, "TtkDonemKar", f"HESAP!$B${t + 5}")
    ad_ekle(wb, "TtkDonemZarar", f"HESAP!$B${t + 6}")
    ad_ekle(wb, "TtkOzsermaye", f"HESAP!$B${t + 7}")
    ad_ekle(wb, "TtkKayipTutari", f"HESAP!$B${t + 8}")
    ad_ekle(wb, "TtkKayipOrani", f"HESAP!$B${t + 9}")
    ad_ekle(wb, "TtkUcteBirEsikTutar", f"HESAP!$B${t + 10}")
    ad_ekle(wb, "TtkUcteIkiEsikTutar", f"HESAP!$B${t + 11}")
    ad_ekle(wb, "TtkToplamBorc", f"HESAP!$B${t + 12}")
    ad_ekle(wb, "TtkToplamAktif", f"HESAP!$B${t + 13}")
    ad_ekle(wb, "TtkBorcaBatik", f"HESAP!$B${t + 14}")
    ad_ekle(wb, "TtkOzKaynakAktif", f"HESAP!$B${t + 15}")
    ad_ekle(wb, "TtkIlkSonFark", f"HESAP!$B${t + 16}")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "TtkUcteBirEsik": 11, "TtkUcteIkiEsik": 12, "TtkTahminCarpan": 13,
        "TtkIyiSenaryo": 14, "TtkBazSenaryo": 15, "TtkKotuSenaryo": 16,
        "TtkTornadoZarar": 17, "TtkKaliteIyi": 18, "TtkKaliteOrta": 19,
        "TtkAksiyon1_ACIKLAMA": 20, "TtkAksiyon2_ACIKLAMA": 21,
        "TtkAksiyon3_ACIKLAMA": 22, "TtkP90Oran": 23, "TtkAzamiBilancoTutari": 24,
        "TtkDonemGirisKolon": 38, "TtkBorcGirisKolon": 39,
    }
    for i in range(1, 13):
        ayar_satirlari[f"TtkPanelSatir{i}"] = 25 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "TtkToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "TtkDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1OzKaynakTrend": 6, "modulT1OzKaynakTrendYorum": 6,
        "modulT4OzKaynakFark": 7, "modulT4OzKaynakFarkYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2Tornado": 9, "modulO2TornadoYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1OzKaynakTahmin": 13, "modulI1OzKaynakTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeBorcTurleri", "LISTELER!$A$7:$A$12")
    ad_ekle(wb, "ListeAylar", "LISTELER!$C$7:$C$18")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "SirketOzKaynagiEridiMiTtkSermayeTamamlama.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
