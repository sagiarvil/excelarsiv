"""
KKEG ve Finansman Gider Kısıtlaması Vergi Savunma Seti — üretim betiği.
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

URUN_AD = "KKEG ve Finansman Gider Kısıtlaması Vergi Savunma Seti"
SURUM = "1.0.0"
RENK = "7A3060"

KAYNAK_TURLERI = ["Banka Kredisi", "Ortak Borcu", "Tahvil", "Finansal Kiralama", "Tedarikçi Borcu"]
EVET_HAYIR = ["Evet", "Hayır"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Dönem öz kaynak ve yabancı kaynak verilerinizi girin; finansman gider "
                "kısıtlamasına göre KKEG tutarını, vergi etkisini ve savunma senaryolarını "
                "otomatik hesaplayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Yıllık ortalama öz kaynak ve yabancı kaynak karşılaştırması (aşan kısım tespiti)",
        "Kapsamdaki finansman gideri üzerinden KKEG tutarı ve kurumlar vergisi etkisi",
        "Kredi türü bazında gider ayrıştırması ve yoğunlaşma riski",
        "Öz kaynak artırımı / borç azaltımı senaryolarında vergi yükü karşılaştırması",
        "UYGUN / İNCELE / DURDUR karar kapısı ve yönetici raporu",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Finansman gideri yüksek olan kurumlar vergisi mükellefleri",
        "Ortaklarından borçlanan ve örtülü sermaye riski taşıyan şirketler",
        "Vergi incelemesine karşı KKEG savunma dosyası hazırlayan mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. DONEM_KAYNAK sayfasına aylık öz kaynak/yabancı kaynak ve gider verilerinizi girin. "
                 "3. KREDI_LISTESI sayfasına kredilerinizi ekleyin. "
                 "4. PANO ve RAPOR sayfalarından KKEG tutarını ve kararı izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Aylık kaynak verilerini girin",
         "DONEM_KAYNAK sayfasındaki sarı hücrelere her ayın öz kaynak, yabancı kaynak ve "
         "finansman gideri değerlerini girin. Ortalamalar otomatik hesaplanır."),
        ("Adım 2 — Kredi listesini girin",
         "KREDI_LISTESI sayfasındaki sarı hücrelere kredilerinizi ve kapsam dışı işaretlerini girin."),
        ("Adım 3 — KKEG hesabını izleyin",
         "HESAP sayfası aşan kısma isabet eden gideri, KKEG tutarını ve kurumlar vergisi etkisini üretir."),
        ("Adım 4 — Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    hatalar = [
        "Öz kaynağı dönem sonu tek değerle girmek (yıllık ortalama bozulur)",
        "Kapsam dışı giderleri ayrıştırmamak (KKEG tutarı şişer)",
        "Kredi listesinde kapsam dışı bayrağını boş bırakmak",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; kurumlar vergisi beyannamesi veya mali "
                 "müşavir görüşü yerine geçmez.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def donem_kaynak(ws):
    sayfa_hazirla(ws, "DONEM_KAYNAK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Dönem Kaynak ve Finansman Gideri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere her ayın öz kaynak, yabancı kaynak ve finansman gideri "
                "değerlerini girin. Ortalamalar ve KKEG katkısı otomatik hesaplanır.",
      yazi=GRİ, kaydir=True)
    sutunlar = ["Ay", "Öz Kaynak (₺)", "Yabancı Kaynak (₺)", "Finansman Gideri (₺)",
                "Kur Farkı Gideri (₺)", "Kapsam Dışı Gider (₺)", "Kapsamdaki Gider (₺)",
                "Yabancı-Öz Kaynak Farkı (₺)"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Kapsamdaki Gider (₺)": '=IF(tblDonem[[#This Row],[Ay]]="","",MAX(0,tblDonem[[#This Row],[Finansman Gideri (₺)]]+tblDonem[[#This Row],[Kur Farkı Gideri (₺)]]-tblDonem[[#This Row],[Kapsam Dışı Gider (₺)]]))',
        "Yabancı-Öz Kaynak Farkı (₺)": '=IF(tblDonem[[#This Row],[Ay]]="","",MAX(0,tblDonem[[#This Row],[Yabancı Kaynak (₺)]]-tblDonem[[#This Row],[Öz Kaynak (₺)]]))',
    }
    tablo_ekle(ws, "tblDonem", "A5:H1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeAylar", "A6:A1004",
              baslik="Ay", mesaj="Listeden dönem ayını seçin (ör. 2026-01).",
              hata_baslik="Geçersiz ay", hata_mesaj="Listede olmayan ay giremezsiniz.")
    dogrulama(ws, "decimal", "0", "B6:B1004",
              baslik="Öz Kaynak", mesaj="Ay sonu öz kaynak tutarını TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Yabancı Kaynak", mesaj="Ay sonu yabancı kaynak tutarını TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000000")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Finansman Gideri", mesaj="Aya ait faiz, komisyon ve vade farkı giderlerini TL girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Kur Farkı Gideri", mesaj="Aya ait kur farkı giderlerini TL girin (yoksa 0).",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Kapsam Dışı Gider", mesaj="Kısıtlama kapsamı dışındaki giderleri TL girin (yoksa 0).",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        for kolon in range(2, 9):
            ws.cell(row=satir, column=kolon).number_format = TL
    ornek = [
        ("2026-01", 45000000, 42000000, 350000, 12000, 0),
        ("2026-02", 45200000, 43500000, 380000, 15000, 0),
        ("2026-03", 45600000, 45100000, 410000, 18000, 0),
        ("2026-04", 46000000, 46800000, 445000, 21000, 0),
        ("2026-05", 46200000, 48500000, 480000, 25000, 0),
        ("2026-06", 46500000, 50200000, 515000, 28000, 0),
        ("2026-07", 46800000, 51900000, 550000, 32000, 0),
        ("2026-08", 47100000, 53600000, 585000, 35000, 0),
        ("2026-09", 47400000, 55300000, 620000, 38000, 0),
        ("2026-10", 47700000, 57000000, 655000, 42000, 0),
        ("2026-11", 48000000, 58700000, 690000, 45000, 0),
        ("2026-12", 48300000, 60400000, 725000, 48000, 0),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            ws.cell(row=i, column=k).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 34, "B": 18, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 22})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def kredi_listesi(ws):
    sayfa_hazirla(ws, "KREDI_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kredi Listesi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere dönem kredilerinizi girin. Kapsam dışı işaretli kredilerin "
                "gideri KKEG hesabına katılmaz; yoğunlaşma ve anomali analizleri bu listeden beslenir.",
      yazi=GRİ, kaydir=True)
    sutunlar = ["Kredi Numarası", "Kaynak Türü", "Kullanım Tarihi", "Vade Tarihi", "Kredi Tutarı (₺)",
                "Yıllık Faiz Oranı", "Dönem Finansman Gideri (₺)", "Kapsam Dışı mı?", "Açıklama",
                "Vade Günü", "Kapsamdaki Gider (₺)"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Vade Günü": '=IF(tblKredi[[#This Row],[Kredi Numarası]]="","",IFERROR(tblKredi[[#This Row],[Vade Tarihi]]-tblKredi[[#This Row],[Kullanım Tarihi]],0))',
        "Kapsamdaki Gider (₺)": '=IF(tblKredi[[#This Row],[Kredi Numarası]]="","",IF(tblKredi[[#This Row],[Kapsam Dışı mı?]]="Evet",0,tblKredi[[#This Row],[Dönem Finansman Gideri (₺)]]))',
    }
    tablo_ekle(ws, "tblKredi", "A5:K1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeKaynakTurleri", "B6:B1004",
              baslik="Kaynak Türü", mesaj="Listeden kaynak türünü seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "list", "ListeEvetHayir", "H6:H1004",
              baslik="Kapsam Dışı mı?", mesaj="Evet seçerseniz bu kredinin gideri KKEG kapsamına girmez.",
              hata_baslik="Geçersiz seçim", hata_mesaj="Evet veya Hayır girin.")
    dogrulama(ws, "date", "01.01.2024", "C6:C1004",
              baslik="Kullanım Tarihi", mesaj="GG.AA.YYYY biçiminde kredi kullanım tarihi girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "date", "01.01.2024", "D6:D1004",
              baslik="Vade Tarihi", mesaj="GG.AA.YYYY biçiminde kredi vade tarihi girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2030 aralığında olmalı.",
              isaret="between", f2="31.12.2030")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Kredi Tutarı", mesaj="Kredi tutarını TL olarak girin.",
              hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Faiz Oranı", mesaj="Yıllık faiz oranını ondalık girin (ör. %25 için 0,25).",
              hata_baslik="Geçersiz oran", hata_mesaj="0 ile 2 arasında olmalı.",
              isaret="between", f2="2")
    dogrulama(ws, "decimal", "0", "G6:G1004",
              baslik="Dönem Gideri", mesaj="Bu krediye ait dönem finansman giderini TL girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = TARİH
        ws.cell(row=satir, column=5).number_format = TL
        ws.cell(row=satir, column=6).number_format = YÜZDE
        ws.cell(row=satir, column=7).number_format = TL
        ws.cell(row=satir, column=10).number_format = CATI
        ws.cell(row=satir, column=11).number_format = TL
    ornek = [
        ("K-1001", "Banka Kredisi", "2026-01-15", "2027-01-15", 10000000, 0.30, 250000, "Hayır", "İşletme kredisi"),
        ("K-1002", "Ortak Borcu", "2026-02-01", "2027-02-01", 5000000, 0.25, 125000, "Hayır", "Ortak avansı"),
        ("K-1003", "Tahvil", "2026-03-10", "2028-03-10", 8000000, 0.28, 186667, "Hayır", "Kurumsal tahvil"),
        ("K-1004", "Banka Kredisi", "2026-04-05", "2027-04-05", 6000000, 0.32, 160000, "Hayır", "Yatırım kredisi"),
        ("K-1005", "Finansal Kiralama", "2026-05-20", "2028-05-20", 3000000, 0.24, 60000, "Evet", "Kiralama kapsam dışı"),
        ("K-1006", "Ortak Borcu", "2026-06-01", "2027-06-01", 4000000, 0.22, 73333, "Hayır", "Ortak katkısı"),
    ]
    for i, satir in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = satir[0]
        ws.cell(row=i, column=2).value = satir[1]
        ws.cell(row=i, column=3).value = date(*map(int, satir[2].split("-")))
        ws.cell(row=i, column=4).value = date(*map(int, satir[3].split("-")))
        for k in (5, 6, 7, 8, 9):
            ws.cell(row=i, column=k).value = satir[k - 1]
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 18, "C": 14, "D": 14, "E": 18, "F": 14, "G": 22, "H": 14, "I": 22, "J": 10, "K": 20})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "KKEG ve Finansman Gider Kısıtlaması Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet, dönem kaynak ve kredi tablolarından otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kaynak Türü", "Kredi Adedi", "Toplam Kredi (₺)", "Toplam Gider (₺)",
                "Kapsamdaki Gider (₺)", "Gider Payı"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(KAYNAK_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=COUNTIFS(tblKredi[Kaynak Türü],$A{r})', sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblKredi[Kredi Tutarı (₺)],tblKredi[Kaynak Türü],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblKredi[Dönem Finansman Gideri (₺)],tblKredi[Kaynak Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblKredi[Kapsamdaki Gider (₺)],tblKredi[Kaynak Türü],$A{r})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=IF($E${11}=0,0,E{r}/$E${11})', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        for k in range(2, 7):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: {tur} kredi metriği | "
                       f"Neden önemli: KKEG hesabını gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    t = 6 + len(KAYNAK_TURLERI)
    h(ws, t, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    for k in range(2, 6):
        h(ws, t, k, f"=SUM({chr(64+k)}6:{chr(64+k)}{t-1})", sayi=(TL if k in (3, 4, 5) else CATI),
          yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t, 6, "=SUM(F6:F10)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 1, 1, "Yıllık Ortalama Öz Kaynak", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 1, 2, "=IFERROR(AVERAGE(tblDonem[Öz Kaynak (₺)]),0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 2, 1, "Yıllık Ortalama Yabancı Kaynak", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 2, 2, "=IFERROR(AVERAGE(tblDonem[Yabancı Kaynak (₺)]),0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 3, 1, "Yabancı-Öz Kaynak Farkı (Aşan Kısım)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 3, 2, f"=MAX(0,B{t + 2}-B{t + 1})", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 4, 1, "Fark Oranı (Aşan Kısım / Yabancı Kaynak)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 4, 2, f"=IF(B{t + 2}=0,0,B{t + 3}/B{t + 2})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 5, 1, "Toplam Kapsamdaki Gider", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 5, 2, "=SUM(tblDonem[Kapsamdaki Gider (₺)])", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 6, 1, "Muafiyet Sonrası Kapsamdaki Gider", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 6, 2, f"=MAX(0,B{t + 5}-KkegMuafiyetEsik)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 7, 1, "Aşan Kısma İsabet Eden Gider", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 7, 2, f"=IF(B{t + 5}=0,0,ROUND(B{t + 6}*B{t + 4},2))", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 8, 1, "KKEG Tutarı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 8, 2, f"=ROUND(B{t + 7}*KkegOran,2)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 9, 1, "Kurumlar Vergisi Etkisi (Ek Vergi Yükü)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 9, 2, f"=ROUND(B{t + 8}*KkegKvOran,2)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 10, 1, "KKEG'nin Toplam Gidere Oranı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 10, 2, f"=IF(B{t + 5}=0,0,B{t + 8}/B{t + 5})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, t + 12, "Hesap satırları yalnızca dönem kaynak ve kredi tablolarından beslenir; korumalıdır.")
    genislik(ws, {"A": 38, "B": 20, "C": 18, "D": 18, "E": 18, "F": 12})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Yabancı Kaynak Payı",
         '=IFERROR(IF(HESAP!B12+HESAP!B13=0,0,HESAP!B13/(HESAP!B12+HESAP!B13)),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"Yabancı kaynağın toplam kaynağa payı ",TEXT(C6,"0.0%")," — pay yüksekse finansman gider kısıtlaması riski belirginleşir")'),
        ("T4", "İlk-Son Ay Öz Kaynak Farkı",
         '=IFERROR(INDEX(tblDonem[Öz Kaynak (₺)],COUNTA(tblDonem[Ay]))-INDEX(tblDonem[Öz Kaynak (₺)],1),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"İlk ile son ay öz kaynak farkı ",TEXT(C7,"₺ #,##0")," TL — fark öz kaynak güçlenme eğilimini gösterir")'),
        ("O1", "Aylık Gider Sapması (STDEV)",
         "=IFERROR(STDEV.P(tblDonem[Kapsamdaki Gider (₺)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Aylık kapsamdaki giderlerin standart sapması ",TEXT(C8,"₺ #,##0")," TL — sapma yüksekse gider yapısı dengesizdir")'),
        ("O2", "Tornado: En Büyük Etki",
         "=MAX(SENARYO_DUYARLILIK!B12:SENARYO_DUYARLILIK!B13)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"KKEG tutarını en çok etkileyen değişkenin etkisi ",TEXT(C9,"₺ #,##0")," TL — öncelik bu değişkenin izlenmesinde")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C9",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kötümser senaryo KKEG farkı ",TEXT(C10,"₺ #,##0")," TL — belirsizlik aralığını gösterir")'),
        ("O6", "Kaynak Türü Yoğunlaşması (HHI)",
         "=IFERROR(SUMPRODUCT((HESAP!$C$6:$C$10/SUM(HESAP!$C$6:$C$10))^2),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kredi yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça kredi tek kaynak türünde toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=KkegKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Kapsamdaki Gider Tahmini",
         "=PANO!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem kapsamdaki gider tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Kredi Tutarı Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblKredi[Kredi Tutarı (₺)],KkegP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kredi tutarlarının %90 ı ",TEXT(C14,"₺ #,##0")," TL altında gerçekleşir — üst sınır nakit planı için kullanılabilir")'),
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
        ("Boş ay satırı var mı?", "=IF(COUNTBLANK(tblDonem[Ay])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Boş öz kaynak var mı?", "=IF(COUNTBLANK(tblDonem[Öz Kaynak (₺)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Negatif değer var mı?", "=IF(COUNTIF(tblDonem[Öz Kaynak (₺)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Yabancı kaynak sıfır mı?", "=IF(COUNTIF(tblDonem[Yabancı Kaynak (₺)],\"=0\")>0,\"SIFIR VAR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Kapsam dışı boş mu?", "=IF(COUNTIF(tblKredi[Kapsam Dışı mı?],\"\")>0,\"BOŞ VAR\",\"TEMİZ\")", NORMAL, KRITIK),
        ("Fark oranı eşikte mi?", "=IF(HESAP!B15>KkegFarkEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("KKEG eşiği aşılıyor mu?", "=IF(HESAP!B19>KkegEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Kapsamdaki gider eşikte mi?", "=IF(HESAP!B16>KkegGiderEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Boş ve negatif girişlerin oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblDonem[Ay])*KkegDonemGirisKolon+COUNTA(tblKredi[Kredi Numarası])*KkegKrediGirisKolon",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblDonem[Ay])+COUNTA(tblDonem[Öz Kaynak (₺)])+COUNTA(tblDonem[Yabancı Kaynak (₺)])+"
                 "COUNTA(tblDonem[Finansman Gideri (₺)])+COUNTA(tblDonem[Kur Farkı Gideri (₺)])+"
                 "COUNTA(tblDonem[Kapsam Dışı Gider (₺)])+COUNTA(tblKredi[Kredi Numarası])+"
                 "COUNTA(tblKredi[Kaynak Türü])+COUNTA(tblKredi[Kredi Tutarı (₺)])+"
                 "COUNTA(tblKredi[Yıllık Faiz Oranı])+COUNTA(tblKredi[Dönem Finansman Gideri (₺)])+"
                 "COUNTA(tblKredi[Kapsam Dışı mı?])+COUNTA(tblKredi[Açıklama])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(KkegToplamGiris=0,0,KkegDoluGiris/KkegToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=KkegKaliteIyi,\"YÜKSEK\",IF(B19>=KkegKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 55})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — KKEG ve Finansman Gider Kısıtlaması", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Fark oranı, KKEG tutarı ve kapsamdaki gider eşiklerine göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "KKEG Tutarı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B19", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Yıllık Ortalama Yabancı Kaynak", yazi="333333")
    h(ws, 7, 2, "=HESAP!B13", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Kapsamdaki Gider", yazi="333333")
    h(ws, 8, 2, "=HESAP!B16", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Fark Oranı", yazi="333333")
    h(ws, 9, 2, "=HESAP!B15", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblDonem[Ay])=0,"VERİ YOK",IF(COUNTIF(tblDonem[Öz Kaynak (₺)],"<0")>0,"DURDUR",IF(AND(COUNTA(tblDonem[Ay])>0,HESAP!B12=0),"DURDUR",IF(HESAP!B15>KkegFarkEsik,"İNCELE",IF(HESAP!B19>KkegEsik,"İNCELE",IF(HESAP!B16>KkegGiderEsik,"İNCELE","UYGUN"))))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblDonem[Öz Kaynak (₺)],"<0")>0,"Negatif öz kaynak girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(AND(COUNTA(tblDonem[Ay])>0,HESAP!B12=0),"Öz kaynak ortalaması sıfır; kısıtlama hesabı yapılamaz, veri girişi düzeltilmeli.",'\
                 'IF(HESAP!B15>KkegFarkEsik,"Yabancı kaynak öz kaynağı belirgin aşıyor; finansman gider kısıtlaması riski yüksek.",'\
                 'IF(HESAP!B19>KkegEsik,"KKEG tutarı eşiğin üzerinde; vergi yükü ve savunma dosyası kapsamlı inceleme gerektirir.",'\
                 'IF(HESAP!B16>KkegGiderEsik,"Kapsamdaki gider eşiğin üzerinde; gider ayrıştırması ve belgeleme kontrol edilmeli.",'\
                 '"Karar kapısı uygun: finansman gider kısıtlaması riski düşük, KKEG tutarı kabul edilebilir.")))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=KkegAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki eşiklerden türetilir; "
                     "kesin vergi incelemesi sonucu değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "KKEG Tutarı", "=HESAP!B19", TL),
        (3, "Yabancı Kaynak Ort.", "=HESAP!B13", TL),
        (4, "Kapsamdaki Gider", "=HESAP!B16", TL),
        (5, "Fark Oranı", "=HESAP!B15", YÜZDE),
        (6, "KKEG Gider Oranı", "=HESAP!B21", YÜZDE),
        (7, "Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Toplam Kredi", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!C11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — kaynak türü bazlı
    h(ws, 10, 1, "Kaynak Türü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 2, "Kredi Adedi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 10, 3, "Toplam Gider (₺)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    for i, tur in enumerate(KAYNAK_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i-5}", sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 6, f"=HESAP!E{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 6, "Kapsamdaki Gider (₺)", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 10, 4, "Aylık Kapsamdaki Gider", kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
    for i in range(12):
        h(ws, 11 + i, 4, f"AY{i+1}", yazi=GRİ)
        h(ws, 11 + i, 5, f"=IFERROR(INDEX(tblDonem[Kapsamdaki Gider (₺)],AYARLAR!$B${28 + i}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Kaynak Türü Kredi Adedi ve Gider"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=15), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F6")
    g2 = PieChart()
    g2.title = "Kredi Adedi Dağılımı"
    g2.add_data(Reference(ws, min_col=2, min_row=10, max_row=15), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U6")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Kaynak Türü Kapsamdaki Gider"
    g3.add_data(Reference(ws, min_col=6, min_row=10, max_row=15), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g3.height, g3.width = 8, 13
    ws.add_chart(g3, "F19")
    g4 = LineChart()
    g4.title = "Aylık Kapsamdaki Gider Trendi"
    g4.add_data(Reference(ws, min_col=5, min_row=10, max_row=23), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=11, max_row=22))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U19")
    # Zaman serisi projeksiyonu
    h(ws, 24, 1, "Gelecek 1 Dönem Kapsamdaki Gider Tahmini (Tahmin Aralıklı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, '=ROUND(IFERROR(FORECAST.LINEAR(KkegPanelSatir12,tblDonem[Kapsamdaki Gider (₺)],ROW(tblDonem[Kapsamdaki Gider (₺)])-ROW(tblDonem[[#Headers],[Kapsamdaki Gider (₺)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Ortalama Aylık Gider", yazi="333333")
    h(ws, 25, 2, "=ROUND(IFERROR(AVERAGE(tblDonem[Kapsamdaki Gider (₺)]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Gider Standart Sapma", yazi="333333")
    h(ws, 26, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblDonem[Kapsamdaki Gider (₺)],KkegPanelSatir1):INDEX(tblDonem[Kapsamdaki Gider (₺)],KkegPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 27, 2, "=B25-B26*KkegTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 28, 2, "=B25+B26*KkegTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=24, max_row=28), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=25, max_row=28))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F33")
    g6 = LineChart()
    g6.title = "Kaynak Türü Toplam Gider"
    g6.add_data(Reference(ws, min_col=3, min_row=10, max_row=15), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U33")
    # Senaryo karşılaştırma verisi
    h(ws, 29, 3, "Senaryo KKEG", kalin=True, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser"], 30):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i-23}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo KKEG Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=29, max_row=32), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=30, max_row=32))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F39")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "KKEG ve Yabancı Kaynak Yapısı"
    g8.add_data(Reference(ws, min_col=2, min_row=3, max_row=4), titles_from_data=True)
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U39")
    baski_hazirla(ws, "A1:M56", URUN_AD)
    alt_bant(ws, 56, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 45, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İyimser, baz ve kötümser senaryolarda KKEG tutarı; ardından hangi değişkenin "
                "sonucu en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    senaryolar = [
        ("İYİMSER", "KkegIyiSenaryo"),
        ("BAZ", "KkegBazSenaryo"),
        ("KÖTÜMSER", "KkegKotuSenaryo"),
    ]
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Kapsamdaki Gider Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "KKEG Tutarı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "KKEG Gider Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, carpan) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=IF(SUM(tblDonem[Kapsamdaki Gider (₺)])=0,0,{carpan})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=ROUND(HESAP!B19*{carpan},0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(HESAP!B19=0,0,C{i}/HESAP!B19)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Tornado: KKEG Tutarına Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tornadolar = [
        ("Yabancı Kaynak +%1", "=ROUND(HESAP!B19*KkegTornadoYabanci,0)", "Yabancı Kaynak"),
        ("Kısıtlama Oranı +%1", "=ROUND(HESAP!B19*KkegTornadoOran,0)", "Kısıtlama Oranı"),
    ]
    for i, (ad, form, etiket) in enumerate(tornadolar, 12):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IF(B12>=B13,\"Yabancı Kaynak\",\"Kısıtlama Oranı\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 17, "Senaryo çarpanları AYARLAR'daki parametrelerden türetilir.")
    genislik(ws, {"A": 45, "B": 30, "C": 20, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "KKEG VE FİNANSMAN GİDER KISITLAMASI — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Yıllık Ortalama Öz Kaynak", "=HESAP!B12", TL),
        ("Yıllık Ortalama Yabancı Kaynak", "=HESAP!B13", TL),
        ("Yabancı-Öz Kaynak Farkı", "=HESAP!B14", TL),
        ("Toplam Kapsamdaki Gider", "=HESAP!B16", TL),
        ("KKEG Tutarı", "=HESAP!B19", TL),
        ("Kurumlar Vergisi Etkisi", "=HESAP!B20", TL),
        ("Fark Oranı", "=HESAP!B15", YÜZDE),
        ("Kaynak Yoğunlaşması (HHI)", "=ANALITIK_MOTOR!C11", "0.000"),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        if sayi:
            h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
        else:
            h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "KKEG tutarı, yabancı kaynak öz kaynağı aştığında aşan kısma isabet eden finansman giderinin kısıtlama oranıyla çarpımıdır; kesin sonuç vergi incelemesine bağlıdır.",
        "Kısıtlama oranı, kurumlar vergisi oranı ve muafiyet eşiği AYARLAR sayfasından alınır; mevzuat değişikliklerinde güncellenmelidir.",
        "Kapsam dışı giderler, kredi listesindeki kapsam dışı işaretleriyle ayrıştırılır; kur farkı giderleri dönem kaynak tablosundan gelir.",
        "Bu dosya kurumlar vergisi beyannamesi yerine geçmez; mali müşavir görüşü gerektirir.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 24 + i, 1, f"=KkegAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
        ws.merge_cells(start_row=24 + i, start_column=1, end_row=24 + i, end_column=12)
    h(ws, 28, 1, "Bu rapor karar destek amaçlıdır; mali tavsiye niteliği taşımaz.",
      yazi=GRİ, boyut=9, italik=True, kaydir=True)
    baski_hazirla(ws, "A1:L28", URUN_AD)
    genislik(ws, {"A": 45, "B": 40})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", "70AD47", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Örnek Veri (Demo)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Gerçek girişlere başlamadan önce demo veri ile ürünü deneyin. "
                "Bu satırlar örnektir; gerçek verilerle değiştirin.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ay", "Öz Kaynak (₺)", "Yabancı Kaynak (₺)", "Finansman Gideri (₺)",
                "Kur Farkı Gideri (₺)", "Kapsam Dışı Gider (₺)", "Kapsamdaki Gider (₺)",
                "Yabancı-Öz Kaynak Farkı (₺)"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Kapsamdaki Gider (₺)": '=IF(tblOrnekDonem[[#This Row],[Ay]]="","",MAX(0,tblOrnekDonem[[#This Row],[Finansman Gideri (₺)]]+tblOrnekDonem[[#This Row],[Kur Farkı Gideri (₺)]]-tblOrnekDonem[[#This Row],[Kapsam Dışı Gider (₺)]]))',
        "Yabancı-Öz Kaynak Farkı (₺)": '=IF(tblOrnekDonem[[#This Row],[Ay]]="","",MAX(0,tblOrnekDonem[[#This Row],[Yabancı Kaynak (₺)]]-tblOrnekDonem[[#This Row],[Öz Kaynak (₺)]]))',
    }
    tablo_ekle(ws, "tblOrnekDonem", "A5:H1004", sutunlar, formuller)
    ornek = [
        ("2026-01", 45000000, 42000000, 350000, 12000, 0),
        ("2026-02", 45200000, 43500000, 380000, 15000, 0),
        ("2026-03", 45600000, 45100000, 410000, 18000, 0),
        ("2026-04", 46000000, 46800000, 445000, 21000, 0),
        ("2026-05", 46200000, 48500000, 480000, 25000, 0),
        ("2026-06", 46500000, 50200000, 515000, 28000, 0),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            h(ws, i, k, deger)
        for k in range(2, 9):
            ws.cell(row=i, column=k).number_format = TL
    h(ws, 3, 10, "Kredi Listesi (Demo)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sutunlar2 = ["Kredi Numarası", "Kaynak Türü", "Kredi Tutarı (₺)", "Yıllık Faiz Oranı",
                 "Dönem Gideri (₺)", "Kapsam Dışı mı?", "Vade Günü", "Kapsamdaki Gider (₺)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar2, basla=10)
    formuller2 = {
        "Vade Günü": '=IF(tblOrnekKredi[[#This Row],[Kredi Numarası]]="","",365)',
        "Kapsamdaki Gider (₺)": '=IF(tblOrnekKredi[[#This Row],[Kredi Numarası]]="","",IF(tblOrnekKredi[[#This Row],[Kapsam Dışı mı?]]="Evet",0,tblOrnekKredi[[#This Row],[Dönem Gideri (₺)]]))',
    }
    tablo_ekle(ws, "tblOrnekKredi", "J5:R1004", sutunlar2, formuller2)
    ornek2 = [
        ("K-1001", "Banka Kredisi", 10000000, 0.30, 250000, "Hayır", "İşletme kredisi"),
        ("K-1002", "Ortak Borcu", 5000000, 0.25, 125000, "Hayır", "Ortak avansı"),
        ("K-1003", "Tahvil", 8000000, 0.28, 186667, "Hayır", "Kurumsal tahvil"),
        ("K-1004", "Banka Kredisi", 6000000, 0.32, 160000, "Hayır", "Yatırım kredisi"),
        ("K-1005", "Finansal Kiralama", 3000000, 0.24, 60000, "Evet", "Kiralama kapsam dışı"),
        ("K-1006", "Ortak Borcu", 4000000, 0.22, 73333, "Hayır", "Ortak katkısı"),
    ]
    for i, satir in enumerate(ornek2, 6):
        for k, deger in enumerate(satir[:6], 0):
            h(ws, i, 10 + k, deger)
        h(ws, i, 18, satir[6])
        ws.cell(row=i, column=12).number_format = TL
        ws.cell(row=i, column=13).number_format = YÜZDE
        ws.cell(row=i, column=14).number_format = TL
        ws.cell(row=i, column=17).number_format = TL
    genislik(ws, {"A": 12, "B": 18, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 22,
                  "J": 24, "K": 18, "L": 18, "M": 14, "N": 18, "O": 16, "P": 10, "Q": 20, "R": 22})
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
    h(ws, 6, 1, "Kaynak Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(KAYNAK_TURLERI, 7):
        h(ws, i, 1, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Kaynak türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni tür için satır ekleyin")
    h(ws, 6, 3, "Ay", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 7 + i, 3, f"2026-{i+1:02d}", zemin=GIRIS_SARI)
    h(ws, 6, 5, "Kapsam Dışı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
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
        (11, "KkegEsik", 500000, "TL", 0, 10000000000, "KKEG tutarı eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (12, "KkegGiderEsik", 10000000, "TL", 0, 100000000000, "Kapsamdaki gider eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (13, "KkegFarkEsik", 0.30, "Oran", 0, 1, "Yabancı-öz kaynak fark oranı eşiği; üzerinde KKEG riski.", "Varsayım"),
        (14, "KkegOran", 0.25, "Oran", 0, 1, "Finansman gider kısıtlama oranı (2023 sonrası için %25).", "Mevzuat"),
        (15, "KkegKvOran", 0.25, "Oran", 0, 1, "Kurumlar vergisi oranı; KKEG vergi etkisi hesabında pay.", "Mevzuat"),
        (16, "KkegMuafiyetEsik", 0, "TL", 0, 10000000000, "Kısıtlama öncesi muafiyet eşiği; uygulanmayacaksa 0.", "Varsayım"),
        (17, "KkegIyiSenaryo", 1.10, "Oran", 0.5, 2, "İyimser senaryoda kapsamdaki gider bu çarpanla hesaplanır.", "Varsayım"),
        (18, "KkegBazSenaryo", 1.00, "Oran", 0.5, 2, "Baz senaryoda kapsamdaki gider bu çarpanla hesaplanır.", "Varsayım"),
        (19, "KkegKotuSenaryo", 0.90, "Oran", 0.5, 2, "Kötümser senaryoda kapsamdaki gider bu çarpanla hesaplanır.", "Varsayım"),
        (20, "KkegTornadoYabanci", 0.01, "Oran", 0, 0.2, "Yabancı kaynağın %1 artışının KKEG tutarına etkisi.", "Varsayım"),
        (21, "KkegTornadoOran", 0.01, "Oran", 0, 0.2, "Kısıtlama oranının %1 artışının KKEG tutarına etkisi.", "Varsayım"),
        (22, "KkegTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (23, "KkegKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (24, "KkegKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (25, "KkegAksiyon1_ACIKLAMA",
         "Fark oranı eşikte ise ortak sermaye artırımı veya borç azaltımıyla yabancı kaynak/öz kaynak dengesini kurun.", "Metin", "", "",
         "İNCELE kararında gösterilen ilk öneri.", "Varsayım"),
        (26, "KkegAksiyon2_ACIKLAMA",
         "KKEG tutarı eşiği aşılıyorsa gider ayrıştırmasını, kur farkı ve kapsam dışı işaretlerini mali müşavirle doğrulayın.", "Metin", "", "",
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (27, "KkegAksiyon3_ACIKLAMA",
         "Kapsamdaki gider eşikte ise finansmanın öz kaynakla yeniden yapılandırılması veya oran uyarlaması için senaryo sayfasını kullanın.", "Metin", "", "",
         "İNCELE kararında gösterilen üçüncü öneri.", "Varsayım"),
        (28, "KkegP90Oran", 0.9, "Oran", 0.5, 1, "Kredi tutarı dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
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
        satir = 28 + i
        h(ws, satir, 1, f"KkegPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki aylık gider grafiği serisinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki aylık grafiği besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    for satir, anahtar, deger, aciklama in [
        (41, "KkegDonemGirisKolon", 6, "DÖNEM_KAYNAK tablosunda satır başına manuel giriş kolonu sayısı."),
        (42, "KkegKrediGirisKolon", 9, "KREDI_LISTESI tablosunda satır başına manuel giriş kolonu sayısı."),
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
         "DONEM_KAYNAK sayfasında sarı hücrelere aylık öz kaynak, yabancı kaynak ve finansman "
         "gideri verilerinizi girin; KREDI_LISTESI sayfasına kredilerinizi ve kapsam dışı işaretlerini ekleyin."),
        ("Hesap katmanı",
         "HESAP sayfası yıllık ortalamaları, aşan kısma isabet eden gideri, KKEG tutarını ve "
         "kurumlar vergisi etkisini otomatik özetler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki kısıtlama oranı, kurumlar vergisi oranı ve eşikleri mali "
         "müşavirinize göre güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; kurumlar vergisi beyannamesi ve mali "
                 "müşavir onayı yerine geçmez.", yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "KKEG Tutarı Eşiği ve Kapsamdaki Gider Eşiği",
        "Yabancı-Öz Kaynak Fark Oranı Risk Eşiği",
        "Muafiyet Eşiği (mevzuatta yoksa 0 bırakılır)",
        "İyimser, Baz ve Kötümser Senaryo Çarpanları",
        "Tornado Yabancı Kaynak ve Kısıtlama Oranı Etkileri",
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

    ws = wb["DONEM_KAYNAK"]
    for harf in ("B", "C", "D", "E", "F"):
        ws.conditional_formatting.add(f"{harf}6:{harf}1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("H6:H1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))

    ws = wb["KREDI_LISTESI"]
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("F6:F1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("H6:H1004", CellIsRule(operator="equal", formula=['"Evet"'],
                                 font=f(amber), fill=d("FFF2CC")))
    ws.conditional_formatting.add("K6:K1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("K6:K1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))

    ws = wb["HESAP"]
    for harf in ("C", "D", "E", "F"):
        ws.conditional_formatting.add(f"{harf}6:{harf}11",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B14:B14", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B15:B15", CellIsRule(operator="greaterThan", formula=["0.3"], font=f(amber)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="greaterThan", formula=["500000"], font=f(amber)))
    ws.conditional_formatting.add("B20:B20", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "SIFIR VAR", "EKSİK VAR", "AŞIYOR", "YÜKSEK"):
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
    ws.conditional_formatting.add("B11:B11", CellIsRule(operator="equal", formula=['"VERİ YOK"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B8:B8", CellIsRule(operator="greaterThan", formula=["10000000"], font=f(amber)))

    ws = wb["PANO"]
    for harf in ("B", "C", "D", "E"):
        ws.conditional_formatting.add(f"{harf}4:{harf}4",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B4:B4", CellIsRule(operator="greaterThan", formula=["500000"], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("F11:F15", CellIsRule(operator="greaterThan", formula=["10000000"], font=f(amber)))
    ws.conditional_formatting.add("E11:E22", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D7:D9", CellIsRule(operator="lessThan", formula=["1"], font=f(kirmizi)))
    ws.conditional_formatting.add("D7:D9", CellIsRule(operator="greaterThan", formula=["1.5"], font=f(amber)))
    ws.conditional_formatting.add("B12:B13", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["ORNEK_VERI"]
    for harf in ("B", "C", "D", "E", "F", "L", "N", "Q"):
        ws.conditional_formatting.add(f"{harf}6:{harf}1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("Q6:Q1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()
    kapak(ws)
    sayfalar = [
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("DONEM_KAYNAK", donem_kaynak),
        ("KREDI_LISTESI", kredi_listesi),
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
    t = 6 + len(KAYNAK_TURLERI)
    ad_ekle(wb, "OrtalamaOzKaynak", f"HESAP!$B${t + 1}")
    ad_ekle(wb, "OrtalamaYabanciKaynak", f"HESAP!$B${t + 2}")
    ad_ekle(wb, "YabanciOzKaynakFarki", f"HESAP!$B${t + 3}")
    ad_ekle(wb, "FarkOrani", f"HESAP!$B${t + 4}")
    ad_ekle(wb, "ToplamKapsamGider", f"HESAP!$B${t + 5}")
    ad_ekle(wb, "MuafiyetSonrasiGider", f"HESAP!$B${t + 6}")
    ad_ekle(wb, "AsanKismaIsabetGider", f"HESAP!$B${t + 7}")
    ad_ekle(wb, "KkegTutari", f"HESAP!$B${t + 8}")
    ad_ekle(wb, "KvEtkisi", f"HESAP!$B${t + 9}")
    ad_ekle(wb, "KkegGiderOrani", f"HESAP!$B${t + 10}")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "KkegEsik": 11, "KkegGiderEsik": 12, "KkegFarkEsik": 13, "KkegOran": 14,
        "KkegKvOran": 15, "KkegMuafiyetEsik": 16, "KkegIyiSenaryo": 17,
        "KkegBazSenaryo": 18, "KkegKotuSenaryo": 19, "KkegTornadoYabanci": 20,
        "KkegTornadoOran": 21, "KkegTahminCarpan": 22, "KkegKaliteIyi": 23,
        "KkegKaliteOrta": 24, "KkegAksiyon1_ACIKLAMA": 25, "KkegAksiyon2_ACIKLAMA": 26,
        "KkegAksiyon3_ACIKLAMA": 27, "KkegP90Oran": 28,
        "KkegDonemGirisKolon": 41, "KkegKrediGirisKolon": 42,
    }
    for i in range(1, 13):
        ayar_satirlari[f"KkegPanelSatir{i}"] = 28 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "KkegToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "KkegDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1YabanciPay": 6, "modulT1YabanciPayYorum": 6,
        "modulT4OzKaynakFark": 7, "modulT4OzKaynakFarkYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2Tornado": 9, "modulO2TornadoYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1GiderTahmin": 13, "modulI1GiderTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeKaynakTurleri", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeAylar", "LISTELER!$C$7:$C$18")
    ad_ekle(wb, "ListeEvetHayir", "LISTELER!$E$7:$E$8")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "KkegVeFinansmanGiderKisitlamasiVergiSavunmaSeti.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
