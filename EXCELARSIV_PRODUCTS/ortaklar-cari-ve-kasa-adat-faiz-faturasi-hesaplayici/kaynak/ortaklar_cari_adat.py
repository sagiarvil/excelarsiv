"""
Ortaklar Cari & Kasa Adat Faiz Faturası Hesaplayıcı — üretim betiği.
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

URUN_AD = "Ortaklar Cari & Kasa Adat Faiz Faturası Hesaplayıcı"
SURUM = "1.0.0"
RENK = "7A5C3E"

ISLEM_TURLERI = ["Alacak", "Borç"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Ortak cari hareketlerini ve kasa bakiyelerini girin; adat üzerinden dönem "
                "faizini, fatura tutarını ve örtülü sermaye riskini otomatik hesaplayın.",
      kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Ortak cari tablosunda her işlemden sonra kümülatif bakiye ve adat",
        "Gün bazlı adat faizi ve dönem sonu faiz faturası tutarı",
        "Kasa gün sonu bakiyeleri üzerinden ayrı adat ve faiz hesabı",
        "Ortak cari bakiyesinin örtülü sermaye/KKEG eşiğini aşıp aşmadığı",
        "Faiz oranı ve vade duyarlılığı, senaryo karşılaştırması ve tahmin bandı",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Ortaklarıyla cari hesap işleyen limited ve anonim şirketler",
        "Kasa bakiyelerini dönem sonu adat ile hesaplayan mali müşavirler",
        "Faiz faturasını düzenli ve mutabakatlı kesmek isteyen işletmeler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. ORTAK_CARI sayfasına ortak işlemlerini, KASA_ADAT sayfasına kasa "
                 "bakiyelerini girin. 2. Faiz oranını AYARLAR'dan kontrol edin. "
                 "3. PANO ve RAPOR sayfalarından fatura tutarını ve kararı izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Ortak cari işlemlerini girin",
         "ORTAK_CARI sayfasındaki sarı hücrelere işlem tarihi, ortak adı, işlem türü "
         "(Alacak/Borç) ve tutarı girin. Bakiye, gün sayısı ve adat otomatik hesaplanır."),
        ("Adım 2 — Kasa bakiyelerini girin",
         "KASA_ADAT sayfasındaki sarı hücrelere dönem içindeki gün sonu kasa bakiyelerini girin."),
        ("Adım 3 — Faiz oranını kontrol edin",
         "AYARLAR sayfasındaki yıllık adat faiz oranını dönem için uygulanan emsal orana göre güncelleyin."),
        ("Adım 4 — Fatura tutarı ve karar",
         "HESAP ve KARAR sayfaları dönem faizini, fatura tutarını ve UYGUN/İNCELE/DURDUR kararını üretir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    hatalar = [
        "İşlem türünü Alacak/Borç olarak işaretlememek (bakiye yönü bozulur)",
        "Faiz oranını dönemin emsal oranına göre güncellememek",
        "Kasa bakiyesini dönem içinde hiç kaydetmemek (adat sıfır kalır)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; faiz faturası düzenlemeden önce mali "
                 "müşavirinizle mutabakat sağlayın.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def ortak_cari(ws):
    sayfa_hazirla(ws, "ORTAK_CARI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Ortak Cari Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere işlem girin; kayıtları ortak bazında gruplu ve tarih sıralı tutun. "
                "Bakiye ortak adı değişince yeniden başlar; adat otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Ortak Adı", "İşlem Türü", "Tutar (₺)", "İşaretli Tutar",
                "Bakiye (₺)", "Gün Sayısı", "Adat (₺·gün)", "Faiz (₺)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "İşaretli Tutar": '=IF(tblOrtak[[#This Row],[Tarih]]="","",IF(tblOrtak[[#This Row],[İşlem Türü]]="Borç",-tblOrtak[[#This Row],[Tutar (₺)]],tblOrtak[[#This Row],[Tutar (₺)]]))',
        "Bakiye (₺)": '=IF(tblOrtak[[#This Row],[Tarih]]="","",IFERROR(IF(INDEX(tblOrtak[Ortak Adı],ROW(tblOrtak[[#This Row],[Tarih]])-ROW(tblOrtak[[#Headers],[Tarih]])-1)=tblOrtak[[#This Row],[Ortak Adı]],INDEX(tblOrtak[Bakiye (₺)],ROW(tblOrtak[[#This Row],[Tarih]])-ROW(tblOrtak[[#Headers],[Tarih]])-1)+tblOrtak[[#This Row],[İşaretli Tutar]],tblOrtak[[#This Row],[İşaretli Tutar]]),tblOrtak[[#This Row],[İşaretli Tutar]]))',
        "Gün Sayısı": '=IF(tblOrtak[[#This Row],[Tarih]]="","",IFERROR(IF(INDEX(tblOrtak[Tarih],ROW(tblOrtak[[#This Row],[Tarih]])-ROW(tblOrtak[[#Headers],[Tarih]])+1)=0,RaporTarihi,INDEX(tblOrtak[Tarih],ROW(tblOrtak[[#This Row],[Tarih]])-ROW(tblOrtak[[#Headers],[Tarih]])+1))-tblOrtak[[#This Row],[Tarih]],RaporTarihi-tblOrtak[[#This Row],[Tarih]]))',
        "Adat (₺·gün)": '=IF(tblOrtak[[#This Row],[Bakiye (₺)]]="",0,ROUND(tblOrtak[[#This Row],[Bakiye (₺)]]*tblOrtak[[#This Row],[Gün Sayısı]],0))',
        "Faiz (₺)": '=IF(tblOrtak[[#This Row],[Adat (₺·gün)]]="",0,ROUND(tblOrtak[[#This Row],[Adat (₺·gün)]]*AdatFaizOran/AdatYilGunSayisi,2))',
    }
    tablo_ekle(ws, "tblOrtak", "A5:J1004", sutunlar, formuller)
    dogrulama(ws, "date", "01.01.2024", "A6:A1004",
              baslik="Tarih", mesaj="GG.AA.YYYY biçiminde işlem tarihi girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "list", "ListeOrtakAdlari", "B6:B1004",
              baslik="Ortak Adı", mesaj="LISTELER'de tanımlı ortak adını seçin.",
              hata_baslik="Geçersiz ortak", hata_mesaj="Listede olmayan ortak giremezsiniz.")
    dogrulama(ws, "list", "ListeIslemTurleri", "C6:C1004",
              baslik="İşlem Türü", mesaj="Alacak (ortak şirkete koyar) veya Borç (şirket ortağa verir) seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Tutar", mesaj="İşlem tutarını TL olarak girin (pozitif).",
              hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=6).number_format = TL
        ws.cell(row=satir, column=7).number_format = CATI
        ws.cell(row=satir, column=8).number_format = TL
        ws.cell(row=satir, column=9).number_format = TL
    ornek = [
        ("2026-01-05", "Ahmet Yılmaz", "Alacak", 500000, "Sermaye katkısı"),
        ("2026-01-12", "Ayşe Demir", "Alacak", 300000, "Sermaye katkısı"),
        ("2026-02-03", "Ahmet Yılmaz", "Borç", 100000, "Avans çekişi"),
        ("2026-02-18", "Mehmet Kaya", "Alacak", 200000, "Ortak katkısı"),
        ("2026-03-07", "Ayşe Demir", "Borç", 50000, "Avans çekişi"),
        ("2026-03-21", "Ahmet Yılmaz", "Alacak", 150000, "Ortak katkısı"),
    ]
    for i, satir in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = date(*map(int, satir[0].split("-")))
        ws.cell(row=i, column=2).value = satir[1]
        ws.cell(row=i, column=3).value = satir[2]
        ws.cell(row=i, column=4).value = satir[3]
        ws.cell(row=i, column=10).value = satir[4]
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 10])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 18, "C": 12, "D": 16, "E": 18, "F": 18, "G": 12, "H": 18, "I": 16, "J": 22})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def kasa_adat(ws):
    sayfa_hazirla(ws, "KASA_ADAT", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kasa Adat Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere dönem içindeki gün sonu kasa bakiyelerini girin. Adat ve "
                "faiz payı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Açıklama", "Gün Sonu Bakiye (₺)", "Gün Sayısı", "Adat (₺·gün)", "Not"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Gün Sayısı": '=IF(tblKasa[[#This Row],[Tarih]]="","",IFERROR(IF(INDEX(tblKasa[Tarih],ROW(tblKasa[[#This Row],[Tarih]])-ROW(tblKasa[[#Headers],[Tarih]])+1)=0,RaporTarihi,INDEX(tblKasa[Tarih],ROW(tblKasa[[#This Row],[Tarih]])-ROW(tblKasa[[#Headers],[Tarih]])+1))-tblKasa[[#This Row],[Tarih]],RaporTarihi-tblKasa[[#This Row],[Tarih]]))',
        "Adat (₺·gün)": '=IF(tblKasa[[#This Row],[Gün Sonu Bakiye (₺)]]="",0,ROUND(tblKasa[[#This Row],[Gün Sonu Bakiye (₺)]]*tblKasa[[#This Row],[Gün Sayısı]],0))',
    }
    tablo_ekle(ws, "tblKasa", "A5:F1004", sutunlar, formuller)
    dogrulama(ws, "date", "01.01.2024", "A6:A1004",
              baslik="Tarih", mesaj="GG.AA.YYYY biçiminde tarih girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Bakiye", mesaj="Gün sonu kasa bakiyesini TL olarak girin.",
              hata_baslik="Geçersiz bakiye", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = CATI
        ws.cell(row=satir, column=5).number_format = TL
    ornek = [
        ("2026-01-31", "Ocak sonu kasa", 85000, "Dönem sonu"),
        ("2026-02-28", "Şubat sonu kasa", 120000, "Dönem sonu"),
        ("2026-03-31", "Mart sonu kasa", 95000, "Dönem sonu"),
        ("2026-04-30", "Nisan sonu kasa", 140000, "Dönem sonu"),
        ("2026-05-31", "Mayıs sonu kasa", 110000, "Dönem sonu"),
        ("2026-06-30", "Haziran sonu kasa", 160000, "Dönem sonu"),
    ]
    for i, satir in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = date(*map(int, satir[0].split("-")))
        ws.cell(row=i, column=2).value = satir[1]
        ws.cell(row=i, column=3).value = satir[2]
        ws.cell(row=i, column=6).value = satir[3]
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 24, "C": 22, "D": 12, "E": 18, "F": 22})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Adat ve Faiz Hesap Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet, ortak cari ve kasa tablolarından otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Ortak Bazında Özet", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 7, ["Ortak", "Net Bakiye (₺)", "Toplam Adat (₺·gün)", "Faiz Payı (₺)", "Bakiye Payı"])
    for i in range(10):
        r = 8 + i
        h(ws, r, 1, f"=IF(LISTELER!A{r}=\"\",\"\",LISTELER!A{r})", yazi=DIS_REF_YEŞIL)
        h(ws, r, 2, f'=IF(A{r}="","",ROUND(SUMIFS(tblOrtak[İşaretli Tutar],tblOrtak[Ortak Adı],A{r}),2))',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=IF(A{r}="","",ROUND(SUMIFS(tblOrtak[Adat (₺·gün)],tblOrtak[Ortak Adı],A{r}),0))',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=IF(A{r}="","",ROUND(C{r}*AdatFaizOran/AdatYilGunSayisi,2))', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=IF($C$18=0,0,IF(A{r}="","",C{r}/$C$18))', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        for k in range(2, 6):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: Ortak {i + 1} cari metriği | "
                       f"Neden önemli: Ortak bazında faiz ve riski gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    h(ws, 18, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 18, 2, "=SUM(B8:B17)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 3, "=SUM(C8:C17)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 4, "=SUM(D8:D17)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 5, "=SUM(E8:E17)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Ortak Cari Net Bakiye", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 20, 2, "=B18", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 21, 1, "Ortak Cari Toplam Adat", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 21, 2, "=C18", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 22, 1, "Ortak Cari Faiz Tutarı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 22, 2, "=D18", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 23, 1, "Kasa Toplam Adat", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 23, 2, "=IFERROR(SUM(tblKasa[Adat (₺·gün)]),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 24, 1, "Kasa Adat Faiz Tutarı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, "=ROUND(B23*AdatFaizOran/AdatYilGunSayisi,2)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Brüt Faiz Tutarı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 25, 2, "=B22+B24", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 26, 1, "Fatura Tutarı (KDV Hariç)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 26, 2, "=ROUND(B25,2)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 27, 1, "Örtülü Sermaye Risk Bakiye", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 27, 2, '=IF(B20<0,ROUND(-B20,2),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Örtülü Sermaye Oranı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 28, 2, '=IF(AdatOzKaynak=0,0,B27/AdatOzKaynak)', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 29, 1, "Ortalama Bakiye (Adat/Gün)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 29, 2, "=IF(B30=0,0,ROUND(C18/B30,2))", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 30, 1, "Toplam Gün (Ortak Cari)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 30, 2, "=IFERROR(SUM(tblOrtak[Gün Sayısı]),0)", sayi=CATI, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 32, "Hesap satırları yalnızca ortak cari, kasa ve AYARLAR'daki orandan beslenir; korumalıdır.")
    genislik(ws, {"A": 30, "B": 22, "C": 22, "D": 18, "E": 14})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Dönem Bakiye Değişim Oranı",
         '=IFERROR(IF(INDEX(tblOrtak[Bakiye (₺)],COUNTA(tblOrtak[Ortak Adı]))=0,0,INDEX(tblOrtak[Bakiye (₺)],COUNTA(tblOrtak[Ortak Adı]))/HESAP!$B$8-1),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"Dönem sonu net bakiye ",TEXT(INDEX(tblOrtak[Bakiye (₺)],COUNTA(tblOrtak[Ortak Adı])),"₺ #,##0")," TL — bakiyenin dönem boyunca yönü izlenir")'),
        ("T2", "Ortak Adat Trendi (Son 3 Kayıt Ortalaması)",
         "=IFERROR(AVERAGE(INDEX(tblOrtak[Adat (₺·gün)],MAX(1,COUNTA(tblOrtak[Ortak Adı])-2)):INDEX(tblOrtak[Adat (₺·gün)],COUNTA(tblOrtak[Ortak Adı]))),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Son üç işlemin ortalama adatı ",TEXT(C7,"₺ #,##0")," — trend adatın arttığını/azaldığını gösterir")'),
        ("T4", "Kasa Ortalama Bakiye Farkı",
         "=IFERROR(HESAP!$B$23/MAX(1,COUNTA(tblKasa[Not]))-INDEX(tblKasa[Gün Sonu Bakiye (₺)],1),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kasa ortalama bakiyesi ile ilk kayıt farkı ",TEXT(C8,"₺ #,##0")," — kasa seviyesindeki eğilimi gösterir")'),
        ("O1", "Ortak Adat Sapması (STDEV)",
         "=IFERROR(STDEV.P(tblOrtak[Adat (₺·gün)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İşlem adatlarının standart sapması ",TEXT(C9,"₺ #,##0")," — sapma yüksekse işlem dengesi bozuktur")'),
        ("O2", "Tornado: En Büyük Etki",
         "=MAX(SENARYO_DUYARLILIK!B12:SENARYO_DUYARLILIK!B13)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Faiz tutarını en çok etkileyen değişkenin etkisi ",TEXT(C10,"₺ #,##0")," TL — öncelik bu değişkenin izlenmesinde")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C9",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kötümser senaryo faiz farkı ",TEXT(C11,"₺ #,##0")," TL — belirsizlik aralığını gösterir")'),
        ("O6", "Ortak Yoğunlaşması (HHI)",
         "=IFERROR(SUMPRODUCT((HESAP!$C$8:$C$17/SUM(HESAP!$C$8:$C$17))^2),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortak adat yoğunlaşma endeksi ",TEXT(C12,"0.000")," — 1 e yaklaştıkça adat tek ortakta toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C13,"0")," puan — ",IF(C13>=AdatKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Faiz Tahmini",
         "=PANO!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem faiz tutarı tahmini ",TEXT(C14,"₺ #,##0")," TL — güven bandı PANO sayfasındadır")'),
        ("I2", "Adat Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblOrtak[Adat (₺·gün)],AdatP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İşlem adatlarının %90 ı ",TEXT(C15,"₺ #,##0")," ₺·gün altında gerçekleşir — üst sınır risk planı içindir")'),
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
        ("Boş ortak adı var mı?", '=IF(COUNTBLANK(tblOrtak[Ortak Adı])=0,"TEMİZ","BOŞ VAR")', NORMAL, KRITIK),
        ("Boş tutar var mı?", '=IF(COUNTBLANK(tblOrtak[Tutar (₺)])=0,"TEMİZ","BOŞ VAR")', NORMAL, KRITIK),
        ("Negatif tutar var mı?", '=IF(COUNTIF(tblOrtak[Tutar (₺)],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, KRITIK),
        ("İşlem türü dışı değer var mı?", '=IF(COUNTIF(tblOrtak[İşlem Türü],"Alacak")+COUNTIF(tblOrtak[İşlem Türü],"Borç")=COUNTA(tblOrtak[Ortak Adı]),"TEMİZ","GEÇERSİZ VAR")', NORMAL, KRITIK),
        ("Kasa negatif bakiye var mı?", '=IF(COUNTIF(tblKasa[Gün Sonu Bakiye (₺)],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, KRITIK),
        ("Net bakiye negatif mi?", '=IF(HESAP!B20<0,"RİSK","NORMAL")', KRITIK, NORMAL),
        ("Faiz tutarı eşiği aşılıyor mu?", '=IF(HESAP!B25>AdatFaizEsik,"AŞIYOR","NORMAL")', KRITIK, NORMAL),
        ("Örtülü sermaye oranı eşiği aşılıyor mu?", '=IF(HESAP!B28>AdatOzKaynakEsik,"AŞIYOR","NORMAL")', KRITIK, NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Boş ve geçersiz girişlerin oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=(COUNTA(tblOrtak[Ortak Adı])+COUNTA(tblOrtak[Tutar (₺)]))+"\
                 "(COUNTA(tblKasa[Açıklama])+COUNTA(tblKasa[Gün Sonu Bakiye (₺)]))",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblOrtak[Ortak Adı])+COUNTA(tblOrtak[Tutar (₺)])+COUNTA(tblKasa[Gün Sonu Bakiye (₺)])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(AdatToplamGiris=0,0,AdatDoluGiris/AdatToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=AdatKaliteIyi,\"YÜKSEK\",IF(B19>=AdatKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 55})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Cari/Adat Faiz Faturası", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Net bakiye, faiz tutarı ve örtülü sermaye oranına göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Brüt Faiz Tutarı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B25", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Ortak Cari Net Bakiye", yazi="333333")
    h(ws, 7, 2, "=HESAP!B20", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Fatura Tutarı (KDV Hariç)", yazi="333333")
    h(ws, 8, 2, "=HESAP!B26", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Örtülü Sermaye Oranı", yazi="333333")
    h(ws, 9, 2, "=HESAP!B28", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblOrtak[Ortak Adı])=0,"VERİ YOK",IF(COUNTIF(tblOrtak[Tutar (₺)],"<0")>0,"DURDUR",IF(COUNTIF(tblKasa[Gün Sonu Bakiye (₺)],"<0")>0,"DURDUR",IF(HESAP!B28>AdatOzKaynakEsik,"İNCELE",IF(HESAP!B25>AdatFaizEsik,"İNCELE","UYGUN")))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblOrtak[Tutar (₺)],"<0")>0,"Negatif tutar girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(COUNTIF(tblKasa[Gün Sonu Bakiye (₺)],"<0")>0,"Negatif kasa bakiyesi var; kasa kayıtları düzeltilmeli.",'\
                 'IF(HESAP!B28>AdatOzKaynakEsik,"Örtülü sermaye oranı eşiğin üzerinde; faizin KKEG riski var, mali müşavirle mutabakat şart.",'\
                 'IF(HESAP!B25>AdatFaizEsik,"Brüt faiz tutarı eşiğin üzerinde; fatura düzenlemeden önce belgeler doğrulanmalı.","Karar kapısı uygun: faiz tutarı ve risk göstergeleri hedef aralıkta."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=AdatAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki eşiklerden türetilir; "
                     "faiz faturası düzenleme onayı değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Brüt Faiz Tutarı", "=HESAP!B25", TL),
        (3, "Fatura Tutarı", "=HESAP!B26", TL),
        (4, "Ortak Cari Net Bakiye", "=HESAP!B20", TL),
        (5, "Örtülü Sermaye Oranı", "=HESAP!B28", YÜZDE),
        (6, "Kalite Skoru", "=KONTROLLER!B19", CATI),
        (7, "Kasa Adat Faizi", "=HESAP!B24", TL),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Toplam Adat (Ortak)", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B21", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Ortak Bazında Net Bakiye", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(10):
        h(ws, 11 + i, 1, f"=HESAP!A{8 + i}", yazi=GRİ)
        h(ws, 11 + i, 2, f"=HESAP!B{8 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 11 + i, 3, f"=HESAP!D{8 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 5, "Ortak Adat Payı", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(10):
        h(ws, 11 + i, 5, f"=HESAP!A{8 + i}", yazi=GRİ)
        h(ws, 11 + i, 6, f"=HESAP!E{8 + i}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Ortak Bazında Net Bakiye"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_row=21), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=20))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F6")
    g2 = PieChart()
    g2.title = "Ortak Adat Payı Dağılımı"
    g2.add_data(Reference(ws, min_col=6, min_row=10, max_row=21), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=5, min_row=11, max_row=20))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U6")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Ortak Bazında Faiz Payı"
    g3.add_data(Reference(ws, min_col=3, min_row=10, max_row=21), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=20))
    g3.height, g3.width = 8, 13
    ws.add_chart(g3, "F19")
    h(ws, 22, 1, "Kasa Bakiye Trendi", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(10):
        h(ws, 23 + i, 1, f"=IFERROR(INDEX(tblKasa[Açıklama],AdatPanelSatir{i + 1}),\"\")", yazi=GRİ)
        h(ws, 23 + i, 2, f"=IFERROR(INDEX(tblKasa[Gün Sonu Bakiye (₺)],AdatPanelSatir{i + 1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Kasa Gün Sonu Bakiye Trendi"
    g4.add_data(Reference(ws, min_col=2, min_row=22, max_row=33), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=23, max_row=32))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U19")
    h(ws, 24, 4, "Gelecek Dönem Faiz Tahmini (Tahmin Aralıklı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 5, '=ROUND(IFERROR(_xlfn.FORECAST.LINEAR(AdatPanelSatir12,tblOrtak[Faiz (₺)],ROW(tblOrtak[Faiz (₺)])-ROW(tblOrtak[[#Headers],[Tarih]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 4, "Son N İşlem Ortalama Faiz", yazi="333333")
    h(ws, 25, 5, '=ROUND(IFERROR(AVERAGE(INDEX(tblOrtak[Faiz (₺)],AdatPanelSatir1):INDEX(tblOrtak[Faiz (₺)],AdatPanelSatir12)),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 4, "Son N İşlem Faiz Standart Sapma", yazi="333333")
    h(ws, 26, 5, '=ROUND(IFERROR(STDEV.P(INDEX(tblOrtak[Faiz (₺)],AdatPanelSatir1):INDEX(tblOrtak[Faiz (₺)],AdatPanelSatir12)),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 4, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 27, 5, "=E25-E26*AdatTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 4, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 28, 5, "=E25+E26*AdatTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=5, min_row=24, max_row=28), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=4, min_row=25, max_row=28))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F33")
    h(ws, 29, 3, "Senaryo Brüt Faiz", kalin=True, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser"], 30):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i - 23}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo Brüt Faiz Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=29, max_row=32), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=30, max_row=32))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F39")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Faiz ve Fatura Yapısı"
    g8.add_data(Reference(ws, min_col=2, min_row=3, max_row=6), titles_from_data=True)
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U33")
    # 8. grafik: kasa adat payı dilimleri (dönem sonu bakiyeleri dağılımı)
    h(ws, 34, 1, "Kasa Adat Payı (İlk 6 Kayıt)", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(6):
        h(ws, 35 + i, 1, f"=IFERROR(INDEX(tblKasa[Açıklama],AdatPanelSatir{i + 1}),\"\")", yazi=GRİ)
        h(ws, 35 + i, 2, f"=IFERROR(INDEX(tblKasa[Adat (₺·gün)],AdatPanelSatir{i + 1}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
    g9 = PieChart()
    g9.title = "Kasa Adat Dağılımı (İlk Kayıtlar)"
    g9.add_data(Reference(ws, min_col=2, min_row=34, max_row=41), titles_from_data=True)
    g9.set_categories(Reference(ws, min_col=1, min_row=35, max_row=40))
    g9.height, g9.width = 8, 13
    ws.add_chart(g9, "U39")
    baski_hazirla(ws, "A1:M56", URUN_AD)
    alt_bant(ws, 56, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 45, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İyimser, baz ve kötümser faiz oranı senaryolarında brüt faiz; ardından "
                "hangi değişkenin sonucu en çok etkilediğini gösteren tornado analizi.",
      yazi=GRİ, kaydir=True)
    senaryolar = [
        ("İYİMSER", "AdatIyiSenaryo"),
        ("BAZ", "AdatBazSenaryo"),
        ("KÖTÜMSER", "AdatKotuSenaryo"),
    ]
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Faiz Oranı Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Brüt Faiz", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "Fatura Tutarı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, carpan) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=IF(SUM(tblOrtak[Tutar (₺)])=0,0,{carpan})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=ROUND(HESAP!B25*{carpan},2)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=ROUND(C{i},2)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Tornado: Brüt Faize Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tornadolar = [
        ("Faiz Oranı +%1", "=ROUND(HESAP!B25*AdatTornadoOran,2)", "Faiz Oranı"),
        ("Adat +%1", "=ROUND(HESAP!B25*AdatTornadoAdat,2)", "Adat"),
    ]
    for i, (ad, form, etiket) in enumerate(tornadolar, 12):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IF(B12>=B13,\"Faiz Oranı\",\"Adat\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 17, "Senaryo çarpanları AYARLAR'daki parametrelerden türetilir.")
    genislik(ws, {"A": 45, "B": 30, "C": 20, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "ORTAK CARİ & KASA ADAT FAİZ FATURASI — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Ortak Cari Net Bakiye", "=HESAP!B20", TL),
        ("Ortak Cari Toplam Adat", "=HESAP!B21", TL),
        ("Ortak Cari Faiz Tutarı", "=HESAP!B22", TL),
        ("Kasa Toplam Adat", "=HESAP!B23", TL),
        ("Kasa Adat Faiz Tutarı", "=HESAP!B24", TL),
        ("Brüt Faiz Tutarı", "=HESAP!B25", TL),
        ("Fatura Tutarı (KDV Hariç)", "=HESAP!B26", TL),
        ("Örtülü Sermaye Oranı", "=HESAP!B28", YÜZDE),
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
        "Faiz oranı AYARLAR sayfasından alınır; dönemin emsal faiz oranına göre güncellenmelidir.",
        "Adat = Bakiye × Gün; gün sayısı işlemler arası süre, son kayıtta rapor tarihine kadardır.",
        "Ortak cari bakiyesinin öz kaynağa oranı eşiği aşarsa örtülü sermaye/KKEG riski İNCELE olarak işaretlenir.",
        "Faiz faturası KDV'ye tabi değildir; fatura tutarı KDV hariç gösterilir.",
        "Bu dosya faiz faturası düzenleme onayı değildir; mali müşavirle mutabakat şarttır.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=AdatAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    baslik_satiri(ws, 6, ["Tarih", "Ortak Adı", "İşlem Türü", "Tutar (₺)", "Açıklama"])
    ornek = [
        ("2026-01-05", "Ahmet Yılmaz", "Alacak", 500000, "Sermaye katkısı"),
        ("2026-01-12", "Ayşe Demir", "Alacak", 300000, "Sermaye katkısı"),
        ("2026-02-03", "Ahmet Yılmaz", "Borç", 100000, "Avans çekişi"),
        ("2026-02-18", "Mehmet Kaya", "Alacak", 200000, "Ortak katkısı"),
        ("2026-03-07", "Ayşe Demir", "Borç", 50000, "Avans çekişi"),
        ("2026-03-21", "Ahmet Yılmaz", "Alacak", 150000, "Ortak katkısı"),
    ]
    for i, satir in enumerate(ornek, 7):
        ws.cell(row=i, column=1).value = date(*map(int, satir[0].split("-")))
        for k, deger in enumerate(satir[1:], 2):
            h(ws, i, k, deger)
        ws.cell(row=i, column=1).number_format = TARİH
        ws.cell(row=i, column=4).number_format = TL
    genislik(ws, {"A": 14, "B": 18, "C": 12, "D": 16, "E": 22})
    alt_bant(ws, 20, "Demo veriler rapor sonuçlarını örnekler; satın alma sonrası silinerek "
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
    h(ws, 6, 1, "Ortak Adları", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    ornek_ortaklar = [
        "Ahmet Yılmaz", "Ayşe Demir", "Mehmet Kaya", "Zeynep Şahin", "Ali Öztürk",
        "Fatma Çelik", "Murat Aydın", "Elif Arslan", "Mustafa Kılıç", "Hüseyin Doğan",
    ]
    for i in range(10):
        r = 7 + i
        h(ws, r, 1, ornek_ortaklar[i], zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{r}", f"Tanım: Ortak adı {i + 1} | Neden önemli: ORTAK_CARI açılır listesini besler | "
                                 "Doğru kullanım: Ortağın adını yazın")
    h(ws, 6, 3, "İşlem Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(ISLEM_TURLERI, 7):
        h(ws, i, 3, tur, zemin=GIRIS_SARI)
    genislik(ws, {"A": 22, "C": 14})


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
        (11, "AdatFaizOran", 0.19, "Oran", 0, 1, "Yıllık adat faiz oranı; dönemin emsal mevduat oranına göre güncellenir.", "Varsayım"),
        (12, "AdatYilGunSayisi", 36500, "Gün", 1, 36500, "Adat faizinde yılın gün bazı (faiz = adat × oran / bu değer).", "Varsayım"),
        (13, "AdatFaizEsik", 250000, "TL", 0, 10000000000, "Brüt faiz tutarı eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (14, "AdatOzKaynak", 10000000, "TL", 0, 100000000000, "Şirket öz kaynağı; örtülü sermaye oranı hesabında payda.", "Varsayım"),
        (15, "AdatOzKaynakEsik", 0.30, "Oran", 0, 1, "Ortak cari bakiyesinin öz kaynağa oranı eşiği; üzerinde KKEG riski.", "Varsayım"),
        (16, "AdatIyiSenaryo", 1.10, "Oran", 0.5, 2, "İyimser senaryoda brüt faiz bu çarpanla hesaplanır.", "Varsayım"),
        (17, "AdatBazSenaryo", 1.00, "Oran", 0.5, 2, "Baz senaryoda brüt faiz bu çarpanla hesaplanır.", "Varsayım"),
        (18, "AdatKotuSenaryo", 0.90, "Oran", 0.5, 2, "Kötümser senaryoda brüt faiz bu çarpanla hesaplanır.", "Varsayım"),
        (19, "AdatTornadoOran", 0.01, "Oran", 0, 0.2, "Faiz oranının %1 artışının brüt faize etkisi.", "Varsayım"),
        (20, "AdatTornadoAdat", 0.01, "Oran", 0, 0.2, "Adatın %1 artışının brüt faize etkisi.", "Varsayım"),
        (21, "AdatTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (22, "AdatKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (23, "AdatKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (24, "AdatAksiyon1_ACIKLAMA",
         "Örtülü sermaye oranı eşikte ise ortak cari bakiyesini azaltın veya faiz oranını emsal ile mutabakatlandırın.", "Metin", "", "",
         "İNCELE kararında gösterilen ilk öneri.", "Varsayım"),
        (25, "AdatAksiyon2_ACIKLAMA",
         "Faiz tutarı eşiği aşılıyorsa fatura düzenlemeden önce belgeleri ve hesapları mali müşavirle doğrulayın.", "Metin", "", "",
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (26, "AdatAksiyon3_ACIKLAMA",
         "Kasa bakiyelerinin dönem boyunca eksiksiz kaydedildiğini doğrulayın; adat hesabı gün sonu bakiyelerine dayanır.", "Metin", "", "",
         "İNCELE kararında gösterilen üçüncü öneri.", "Varsayım"),
        (27, "AdatP90Oran", 0.9, "Oran", 0.5, 1, "Adat dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
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
        satir = 27 + i
        h(ws, satir, 1, f"AdatPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki kasa trendi ve faiz tahmini grafiğinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki grafikleri besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 30, "B": 22, "C": 12, "D": 10, "E": 10, "F": 60, "G": 14, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "ORTAK_CARI sayfasındaki sarı hücrelere ortak işlemlerini, KASA_ADAT sayfasındaki sarı hücrelere "
         "gün sonu kasa bakiyelerini girin. Ortak cari kayıtlarını ortak bazında gruplu ve tarih sıralı "
         "tutun; bakiye, gün sayısı ve adat otomatik hesaplanır."),
        ("Hesap katmanı",
         "HESAP sayfası ortak bazında net bakiye, adat, faiz payı, kasa adat faizi, brüt faiz ve fatura "
         "tutarını üretir; örtülü sermaye oranını izler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki yıllık faiz oranını, öz kaynak ve eşikleri döneme göre güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; faiz faturası düzenlemeden önce mali müşavirle "
                 "mutabakat sağlayın.", yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "Yıllık Adat Faiz Oranı",
        "Brüt Faiz Tutarı Eşiği",
        "Şirket Öz Kaynağı ve Örtülü Sermaye Oranı Eşiği",
        "İyimser, Baz ve Kötümser Senaryo Çarpanları",
        "Tornado Faiz Oranı ve Adat Etkisi Oranları",
        "Tahmin Çarpanı (standart sapma katı)",
        "Kalite Skoru İyi ve Orta Eşikleri",
        "P90 Yüzdelik Oranı",
        "Aksiyon Önerisi Açıklamaları",
    ]
    for j, v in enumerate(varsayimlar, 15):
        h(ws, j, 1, "• " + v, yazi="333333", boyut=9, kaydir=True)
    h(ws, 25, 1, "Bu dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. "
                "Makro ve dış bağlantı yoktur.", yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells(start_row=25, start_column=2, end_row=25, end_column=20)
    genislik(ws, {"A": 45, "B": 60})


def kosullu_bicimlendirme(wb):
    yesil, amber, kirmizi = "1F7A4D", "B7791F", "B3261E"

    def f(renk):
        return Font(name=FONT, color=renk)

    def d(renk):
        return PatternFill("solid", start_color=renk, end_color=renk)

    ws = wb["ORTAK_CARI"]
    ws.conditional_formatting.add("A6:A1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("C6:C1004", CellIsRule(operator="equal", formula=['"Borç"'], font=f(amber)))

    ws = wb["KASA_ADAT"]
    ws.conditional_formatting.add("C6:C1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))

    ws = wb["HESAP"]
    for harf in ("B", "C", "D"):
        ws.conditional_formatting.add(f"{harf}8:{harf}18",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B20:B20", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B25:B25", CellIsRule(operator="greaterThan", formula=["250000"], font=f(amber)))
    ws.conditional_formatting.add("B28:B28", CellIsRule(operator="greaterThan", formula=["0.3"], font=f(kirmizi)))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "GEÇERSİZ VAR", "RİSK", "AŞIYOR"):
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

    ws = wb["PANO"]
    for harf in ("B", "C", "D"):
        ws.conditional_formatting.add(f"{harf}4:{harf}4",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:B21", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E11:E21", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B12:B13", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["ORNEK_VERI"]
    ws.conditional_formatting.add("D7:D20", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()
    kapak(ws)
    sayfalar = [
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("ORTAK_CARI", ortak_cari),
        ("KASA_ADAT", kasa_adat),
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
    ad_ekle(wb, "OrtakCariNetBakiye", "HESAP!$B$20")
    ad_ekle(wb, "OrtakCariToplamAdat", "HESAP!$B$21")
    ad_ekle(wb, "OrtakCariFaiz", "HESAP!$B$22")
    ad_ekle(wb, "KasaToplamAdat", "HESAP!$B$23")
    ad_ekle(wb, "KasaAdatFaiz", "HESAP!$B$24")
    ad_ekle(wb, "BrutFaizTutari", "HESAP!$B$25")
    ad_ekle(wb, "FaturaTutari", "HESAP!$B$26")
    ad_ekle(wb, "OrtuluSermayeOrani", "HESAP!$B$28")
    ad_ekle(wb, "ToplamAdat", "HESAP!$C$18")
    ad_ekle(wb, "AdatToplamGun", "HESAP!$B$30")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "AdatFaizOran": 11, "AdatYilGunSayisi": 12, "AdatFaizEsik": 13, "AdatOzKaynak": 14,
        "AdatOzKaynakEsik": 15, "AdatIyiSenaryo": 16, "AdatBazSenaryo": 17,
        "AdatKotuSenaryo": 18, "AdatTornadoOran": 19, "AdatTornadoAdat": 20,
        "AdatTahminCarpan": 21, "AdatKaliteIyi": 22, "AdatKaliteOrta": 23,
        "AdatAksiyon1_ACIKLAMA": 24, "AdatAksiyon2_ACIKLAMA": 25,
        "AdatAksiyon3_ACIKLAMA": 26, "AdatP90Oran": 27,
    }
    for i in range(1, 13):
        ayar_satirlari[f"AdatPanelSatir{i}"] = 27 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "AdatToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "AdatDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1DonemBakiye": 6, "modulT1DonemBakiyeYorum": 6,
        "modulT2FaizTrend": 7, "modulT2FaizTrendYorum": 7,
        "modulT4DonemFark": 8, "modulT4DonemFarkYorum": 8,
        "modulO1Anomali": 9, "modulO1AnomaliYorum": 9,
        "modulO2Tornado": 10, "modulO2TornadoYorum": 10,
        "modulO3Senaryo": 11, "modulO3SenaryoYorum": 11,
        "modulO6Hhi": 12, "modulO6HhiYorum": 12,
        "modulO8KaliteSkor": 13, "modulO8KaliteSkorYorum": 13,
        "modulI1AdatTahmin": 14, "modulI1AdatTahminYorum": 14,
        "modulI2YuzdelikP90": 15, "modulI2YuzdelikP90Yorum": 15,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeOrtakAdlari", "LISTELER!$A$7:$A$16")
    ad_ekle(wb, "ListeIslemTurleri", "LISTELER!$C$7:$C$8")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "OrtaklarCariVeKasaAdatFaizFaturasiHesaplayici.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
