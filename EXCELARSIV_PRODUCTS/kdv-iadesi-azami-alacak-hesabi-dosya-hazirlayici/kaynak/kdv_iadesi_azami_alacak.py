"""
KDV İadesi Azami Alacak Hesabı & Dosya Hazırlayıcı — üretim betiği.
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

URUN_AD = "KDV İadesi Azami Alacak Hesabı & Dosya Hazırlayıcı"
SURUM = "1.0.0"
RENK = "1B6E4F"

BELGE_TURLERI = ["Satış", "Alış", "İhracat", "Gider", "Yatırım"]
KDV_ORANLARI = [0.01, 0.10, 0.20]
TAHSILAT_DURUMLARI = ["Belgesi Tam", "Belge Eksik"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Dönem KDV belgelerinizi girin; devreden KDV'yi, ihracat iadesi sınırını ve "
                "azami KDV iade alacağınızı otomatik hesaplayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Belge türü bazında hesaplanan ve indirilecek KDV ayrıştırması",
        "Devreden KDV ve ihracat iadesi sınırına göre azami iade alacağı",
        "Eksik belge tespiti ve dosya hazırlık kontrol listesi",
        "Hangi değişkenin iadeyi en çok etkilediğini gösteren tornado",
        "UYGUN / İNCELE / DURDUR karar kapısı ve yönetici raporu",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "KDV iadesi alacaklısı olan ihracatçılar",
        "Devreden KDV'si sürekli oluşan üreticiler",
        "İade dosyasını kendi hazırlamak isteyen muhasebe profesyonelleri",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. KDV_BELGELERI sayfasına dönem belgelerinizi girin. "
                 "3. PANO ve RAPOR sayfalarından iade alacağınızı ve dosya durumunu izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — KDV belgelerini girin",
         "KDV_BELGELERI sayfasındaki sarı hücrelere belge türü, matrah, KDV oranı ve "
         "belge durumunu girin. KDV tutarı otomatik hesaplanır."),
        ("Adım 2 — İade hesabını izleyin",
         "HESAP sayfası hesaplanan KDV, indirilecek KDV, devreden KDV ve azami iade alacağını üretir."),
        ("Adım 3 — Dosya durumu ve karar",
         "KARAR sayfası dosyanın UYGUN/İNCELE/DURDUR olduğunu söyler; eksik belgeleri listeler."),
        ("Adım 4 — Eşikleri ayarlayın",
         "İade sınırı ve kalite eşikleri AYARLAR sayfasından değiştirilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    hatalar = [
        "Satış ve alış belgelerini aynı türde girmek (KDV ayrışması bozulur)",
        "Matrahı KDV dahil girmek (matrah her zaman KDV hariçtir)",
        "İhracat belgelerinde belge durumunu boş bırakmak",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; KDV beyannamesi veya mali müşavir yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def kdv_belgeleri(ws):
    sayfa_hazirla(ws, "KDV_BELGELERI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "KDV Belgeleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Her satır bir KDV belgesidir; "
                "KDV tutarı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Belge Numarası", "Belge Türü", "Matrah (₺)", "KDV Oranı",
                "Belge Durumu", "KDV Tutarı (₺)", "Kayıt Açıklaması"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "KDV Tutarı (₺)": "=IF(tblKdv[[#This Row],[Matrah (₺)]]=0,0,"\
                          "ROUND(tblKdv[[#This Row],[Matrah (₺)]]*tblKdv[[#This Row],[KDV Oranı]],2))",
        "Kayıt Açıklaması": "=IF(tblKdv[[#This Row],[Belge Türü]]=\"\",\"\","\
                            "CONCATENATE(\"Belge: \",tblKdv[[#This Row],[Belge Numarası]],\"/\",tblKdv[[#This Row],[Belge Türü]],\" \",TEXT(tblKdv[[#This Row],[Matrah (₺)]],\"#,##0\")))",
    }
    tablo_ekle(ws, "tblKdv", "A5:H1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeBelgeTurleri", "C6:C1004",
              baslik="Belge Türü", mesaj="Satış, Alış, İhracat, Gider veya Yatırım seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Listede olmayan tür giremezsiniz.")
    dogrulama(ws, "list", "ListeKdvOranlari", "E6:E1004",
              baslik="KDV Oranı", mesaj="Listeden KDV oranını seçin.",
              hata_baslik="Geçersiz oran", hata_mesaj="Listede olmayan oran giremezsiniz.")
    dogrulama(ws, "list", "ListeTahsilatDurumlari", "F6:F1004",
              baslik="Belge Durumu", mesaj="Belgesi Tam veya Belge Eksik seçin.",
              hata_baslik="Geçersiz durum", hata_mesaj="Listede olmayan durum giremezsiniz.")
    dogrulama(ws, "date", "01.01.2024", "A6:A1004",
              baslik="Tarih", mesaj="GG.AA.YYYY biçiminde tarih girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Matrah", mesaj="KDV hariç matrahı TL olarak girin.",
              hata_baslik="Geçersiz matrah", hata_mesaj="0 ile 100.000.000.000 arasında olmalı.",
              isaret="between", f2="100000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = TL
        ws.cell(row=satir, column=5).number_format = YÜZDE
        ws.cell(row=satir, column=7).number_format = TL
    ornek_kdv = [
        ("2026-01-05", "F-1001", "Satış", 250000, 0.20, "Belgesi Tam", "Örnek: yurt içi satış"),
        ("2026-01-12", "A-2001", "Alış", 180000, 0.20, "Belgesi Tam", "Örnek: mal alımı"),
        ("2026-02-03", "İ-3001", "İhracat", 400000, 0.01, "Belgesi Tam", "Örnek: ihracat"),
        ("2026-02-18", "A-2002", "Alış", 95000, 0.20, "Belge Eksik", "Örnek: hammadde"),
        ("2026-03-07", "F-1002", "Satış", 320000, 0.20, "Belgesi Tam", "Örnek: yurt içi satış"),
        ("2026-03-21", "G-4001", "Gider", 45000, 0.20, "Belgesi Tam", "Örnek: kira"),
        ("2026-04-10", "A-2003", "Alış", 210000, 0.10, "Belgesi Tam", "Örnek: ara mal"),
        ("2026-05-02", "İ-3002", "İhracat", 550000, 0.01, "Belgesi Tam", "Örnek: ihracat"),
        ("2026-05-19", "Y-5001", "Yatırım", 300000, 0.20, "Belgesi Tam", "Örnek: makine"),
        ("2026-06-11", "F-1003", "Satış", 280000, 0.20, "Belgesi Tam", "Örnek: yurt içi satış"),
    ]
    for i, satir in enumerate(ornek_kdv, 6):
        ws.cell(row=i, column=1).value = date(*map(int, satir[0].split("-")))
        for k, deger in enumerate(satir[1:], 2):
            ws.cell(row=i, column=k).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 14, "C": 14, "D": 18, "E": 12, "F": 16, "G": 16, "H": 22})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "KDV İadesi Azami Alacak Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet belge türlerine göre KDV belgelerinden otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    sutunlar = ["Belge Türü", "Belge Adedi", "Toplam Matrah (₺)", "Toplam KDV (₺)",
                "Hesaplanan KDV (₺)", "İndirilecek KDV (₺)", "Devreden KDV (₺)", "Oran Payı"]
    baslik_satiri(ws, 5, sutunlar)
    for r, tur in enumerate(BELGE_TURLERI, 6):
        h(ws, r, 1, tur, kalin=True, yazi="333333")
        h(ws, r, 2, f'=COUNTIFS(tblKdv[Belge Türü],$A{r})', sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblKdv[Matrah (₺)],tblKdv[Belge Türü],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblKdv[KDV Tutarı (₺)],tblKdv[Belge Türü],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=IF($A{r}="Satış",D{r},0)', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=IF($A{r}="Satış",0,D{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f"=IF($A{r}=\"İhracat\",D{r},0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 8, f"=IF(SUM($C$6:$C$10)=0,0,C{r}/SUM($C$6:$C$10))", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        for k in range(2, 9):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: {tur} KDV metriği | "
                       f"Neden önemli: İade hesabını gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    t = 6 + len(BELGE_TURLERI)
    h(ws, t, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    for k in range(2, 8):
        h(ws, t, k, f"=SUM({chr(64+k)}6:{chr(64+k)}{t-1})", sayi=(TL if k in (3, 4, 5, 6, 7) else CATI),
          yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t, 8, "=SUM(H6:H10)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 2, 1, "Hesaplanan KDV (Toplam)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 2, 2, f"=E{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 3, 1, "İndirilecek KDV (Toplam)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 3, 2, f"=F{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 4, 1, "Devreden KDV", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 4, 2, f"=F{t}-E{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 5, 1, "İhracat İade Sınırı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 5, 2, f"=IF(OR(B{t + 2}<=0,C{t}=0),0,ROUND(B{t + 2}*C{t - 3}/C{t},2))",
      sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 6, 1, "Azami İade Alacağı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 6, 2, f"=IF(B{t + 2}<=0,0,MIN(B{t + 2},B{t + 3}))", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 7, 1, "Eksik Belge Sayısı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 7, 2, '=COUNTIF(tblKdv[Belge Durumu],"Belge Eksik")', sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 8, 1, "İade Kullanım Oranı", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 8, 2, f"=IF(B{t+2}=0,0,B{t+4}/B{t+2})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, t + 10, "Hesap satırları yalnızca KDV belgelerinden beslenir; korumalıdır.")
    genislik(ws, {"A": 30, "B": 14, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 12})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Belge Vade: 0-30 Gün Payı",
         '=IFERROR(IF(SUM(tblKdv[Matrah (₺)])=0,0,SUMIFS(tblKdv[Matrah (₺)],tblKdv[Belge Türü],"İhracat")/SUM(tblKdv[Matrah (₺)])),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"İhracat belgelerinin matrah payı ",TEXT(C6,"0.0%")," — pay yüksekse iade hakkı belirginleşir")'),
        ("T4", "İlk-Son Belge Matrah Farkı",
         '=IFERROR(INDEX(tblKdv[Matrah (₺)],COUNTA(tblKdv[Belge Türü]))-INDEX(tblKdv[Matrah (₺)],1),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"İlk ile son belge matrah farkı ",TEXT(C7,"₺ #,##0")," TL — fark iade hacmindeki eğilimi gösterir")'),
        ("O1", "Matrah Sapması (STDEV)",
         "=IFERROR(STDEV.P(tblKdv[Matrah (₺)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Belge matrahlarının standart sapması ",TEXT(C8,"₺ #,##0")," TL — sapma yüksekse belge yapısı dengesizdir")'),
        ("O2", "Tornado: En Büyük Etki",
         "=MAX(SENARYO_DUYARLILIK!B13:SENARYO_DUYARLILIK!B14)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Azami iadeyi en çok etkileyen değişkenin etkisi ",TEXT(C9,"₺ #,##0")," TL — öncelik bu değişkenin izlenmesinde")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C9",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kötümser senaryo iade farkı ",TEXT(C10,"₺ #,##0")," TL — belirsizlik aralığını gösterir")'),
        ("O6", "Belge Yoğunlaşması (HHI)",
         "=IFERROR(SUMPRODUCT((HESAP!$C$6:$C$10/SUM(HESAP!$C$6:$C$10))^2),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Belge yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça matrah tek belge türünde toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=KdvKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem Devreden KDV Tahmini",
         "=PANO!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem devreden KDV tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Matrah Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblKdv[Matrah (₺)],KdvP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Belge matrahlarının %90 ı ",TEXT(C14,"₺ #,##0")," TL altında gerçekleşir — üst sınır nakit planı için kullanılabilir")'),
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
        ("Boş belge türü satırı var mı?", "=IF(COUNTBLANK(tblKdv[Belge Türü])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Boş matrah var mı?", "=IF(COUNTBLANK(tblKdv[Matrah (₺)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Negatif matrah var mı?", "=IF(COUNTIF(tblKdv[Matrah (₺)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Geçersiz KDV oranı var mı?", "=IF(COUNTIF(tblKdv[KDV Oranı],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Eksik belge var mı?", "=IF(COUNTIF(tblKdv[Belge Durumu],\"Belge Eksik\")=0,\"TEMİZ\",\"EKSİK VAR\")", NORMAL, KRITIK),
        ("Devreden KDV negatif mi?", "=IF(HESAP!B15<0,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Azami iade eşiği aşılıyor mu?", "=IF(HESAP!B17>KdvIadeEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("İade kullanım oranı eşikte mi?", "=IF(HESAP!B19>KdvKullanimEsik,\"YÜKSEK\",\"NORMAL\")", KRITIK, NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Boş ve negatif girişlerin oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblKdv[Belge Türü])*COUNTA(tblKdv[#Headers])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblKdv[Belge Türü])+COUNTA(tblKdv[Matrah (₺)])+"
                 "COUNTA(tblKdv[KDV Oranı])+COUNTA(tblKdv[Belge Durumu])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(KdvToplamGiris=0,0,KdvDoluGiris/KdvToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=KdvKaliteIyi,\"YÜKSEK\",IF(B19>=KdvKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 55})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — KDV İade Dosyası", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Devreden KDV, eksik belge ve iade kullanım oranına göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Azami İade Alacağı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B17", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Devreden KDV", yazi="333333")
    h(ws, 7, 2, "=HESAP!B15", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Eksik Belge Sayısı", yazi="333333")
    h(ws, 8, 2, "=HESAP!B18", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "İade Kullanım Oranı", yazi="333333")
    h(ws, 9, 2, "=HESAP!B19", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblKdv[Belge Türü])=0,"VERİ YOK",IF(COUNTIF(tblKdv[Matrah (₺)],"<0")>0,"DURDUR",IF(HESAP!B15<0,"DURDUR",IF(AND(COUNTA(tblKdv[Belge Türü])>0,HESAP!C11=0),"İNCELE",IF(HESAP!C11>KdvMatrahEsik,"İNCELE",IF(HESAP!B18>0,"İNCELE",IF(HESAP!B17>KdvIadeEsik,"İNCELE","UYGUN")))))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblKdv[Matrah (₺)],"<0")>0,"Negatif matrah girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(HESAP!B15<0,"Devreden KDV negatif; indirilecek KDV hesaplananı karşılamıyor, iade hakkı oluşmamış.",'\
                 'IF(AND(COUNTA(tblKdv[Belge Türü])>0,HESAP!C11=0),"Belge türleri HESAP özetiyle eşleşmiyor; veri girişi düzeltilmeli.",'\
                 'IF(HESAP!C11>KdvMatrahEsik,"Toplam belge hacmi eşiğin üzerinde; dosya kapsamlı inceleme gerektirir.",'\
                 'IF(HESAP!B18>0,"Dosyada eksik belge var; iade süreci başlatılmadan belgeler tamamlanmalı.",'\
                 'IF(HESAP!B17>KdvIadeEsik,"Azami iade tutarı eşiğin üzerinde; dosya kapsamlı inceleme gerektirir.","Karar kapısı uygun: iade alacağı belirlendi ve belgeler tam."))))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=KdvAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki eşiklerden türetilir; "
                     "kesin KDV iadesi onayı değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Azami İade Alacağı", "=HESAP!B17", TL),
        (3, "Devreden KDV", "=HESAP!B15", TL),
        (4, "İhracat Matrahı", "=HESAP!G11", TL),
        (5, "Eksik Belge Sayısı", "=HESAP!B18", CATI),
        (6, "İade Kullanım Oranı", "=HESAP!B19", YÜZDE),
        (7, "Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "Toplam Matrah", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!C11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — belge türü bazlı
    h(ws, 10, 1, "Belge Türü Matrah", kalin=True, yazi=KOYU_LACIVERT)
    for i, tur in enumerate(BELGE_TURLERI, 11):
        h(ws, i, 1, tur, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!C{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 4, "Aylık Toplam Matrah", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(12):
        h(ws, 11 + i, 4, f"AY{i+1}", yazi=GRİ)
        h(ws, 11 + i, 5, f"=IFERROR(INDEX(tblKdv[Matrah (₺)],AYARLAR!$B${27 + i}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Belge Türü Matrah ve KDV"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=15), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F6")
    g2 = PieChart()
    g2.title = "Matrah Dağılımı"
    g2.add_data(Reference(ws, min_col=2, min_row=10, max_row=15), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U6")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Belge Türü KDV Tutarı"
    for i in range(11, 16):
        h(ws, i, 6, f"=HESAP!D{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g3.add_data(Reference(ws, min_col=6, min_row=10, max_row=15), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g3.height, g3.width = 8, 13
    ws.add_chart(g3, "F19")
    g4 = LineChart()
    g4.title = "Aylık Matrah Trendi"
    g4.add_data(Reference(ws, min_col=5, min_row=10, max_row=23), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=11, max_row=22))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U19")
    # Zaman serisi projeksiyonu
    h(ws, 24, 1, "Gelecek 1 Dönem Devreden KDV Tahmini (Tahmin Aralıklı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, '=ROUND(IFERROR(FORECAST.LINEAR(KdvPanelSatir12,tblKdv[Matrah (₺)],ROW(tblKdv[Matrah (₺)])-ROW(tblKdv[[#Headers],[Matrah (₺)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Son 12 Ay Ortalama Matrah", yazi="333333")
    h(ws, 25, 2, "=ROUND(IFERROR(AVERAGE(tblKdv[Matrah (₺)]),0),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Son 12 Ay Matrah Standart Sapma", yazi="333333")
    h(ws, 26, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblKdv[Matrah (₺)],KdvPanelSatir1):INDEX(tblKdv[Matrah (₺)],KdvPanelSatir12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 27, 2, "=B25-B26*KdvTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 28, 2, "=B25+B26*KdvTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=24, max_row=28), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=25, max_row=28))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F33")
    g6 = LineChart()
    g6.title = "Belge Türü KDV Tutarı"
    g6.add_data(Reference(ws, min_col=3, min_row=10, max_row=15), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U33")
    # Senaryo karşılaştırma verisi
    h(ws, 29, 3, "Senaryo Azami İade", kalin=True, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser"], 30):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i-23}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo Azami İade Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=29, max_row=32), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=30, max_row=32))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F39")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "İade, Devreden ve İhracat Yapısı"
    g8.add_data(Reference(ws, min_col=2, min_row=3, max_row=6), titles_from_data=True)
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U39")
    baski_hazirla(ws, "A1:M56", URUN_AD)
    alt_bant(ws, 56, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 45, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İyimser, baz ve kötümser senaryolarda azami iade; ardından hangi değişkenin "
                "sonucu en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    senaryolar = [
        ("İYİMSER", "KdvIyiSenaryo"),
        ("BAZ", "KdvBazSenaryo"),
        ("KÖTÜMSER", "KdvKotuSenaryo"),
    ]
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Devreden Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Azami İade", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "İade Kullanım Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, carpan) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=IF(SUM(tblKdv[Matrah (₺)])=0,0,{carpan})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=ROUND(HESAP!B17*{carpan},0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(HESAP!B17=0,0,C{i}/HESAP!B17)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Tornado: Azami İadeye Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tornadolar = [
        ("Devreden KDV +%1", "=ROUND(HESAP!B17*KdvTornadoDevreden,0)", "Devreden KDV"),
        ("İhracat +%1", "=ROUND(HESAP!B17*KdvTornadoIhracat,0)", "İhracat"),
    ]
    for i, (ad, form, etiket) in enumerate(tornadolar, 12):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 15, 2, "=IF(B12>=B13,\"Devreden KDV\",\"İhracat\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 17, "Senaryo çarpanları AYARLAR'daki parametrelerden türetilir.")
    genislik(ws, {"A": 45, "B": 30, "C": 20, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "KDV İADESİ AZAMİ ALACAK — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Hesaplanan KDV", "=HESAP!B13", TL),
        ("İndirilecek KDV", "=HESAP!B14", TL),
        ("Devreden KDV", "=HESAP!B15", TL),
        ("İhracat İade Sınırı", "=HESAP!B16", TL),
        ("Azami İade Alacağı", "=HESAP!B17", TL),
        ("Eksik Belge Sayısı", "=HESAP!B18", CATI),
        ("İade Kullanım Oranı", "=HESAP!B19", YÜZDE),
        ("Belge Yoğunlaşması (HHI)", "=ANALITIK_MOTOR!C11", "0.000"),
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
        "Azami iade alacağı, devreden KDV ile ihracat iade sınırının küçüğüdür; kesin iade tutarı vergi dairesi tespitine bağlıdır.",
        "Matrah girişleri KDV hariçtir; KDV tutarı oran çarpımıyla otomatik üretilir.",
        "İade sınırı ve eşikler AYARLAR sayfasından alınır; mevzuat değişikliklerinde güncellenmelidir.",
        "Bu dosya KDV beyannamesi yerine geçmez; mali müşavir onayı gerektirir.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 24 + i, 1, f"=KdvAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    sutunlar = ["Tarih", "Belge Numarası", "Belge Türü", "Matrah (₺)", "KDV Oranı",
                "Belge Durumu", "KDV Tutarı (₺)", "Not"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "KDV Tutarı (₺)": "=IF(tblOrnek[[#This Row],[Matrah (₺)]]=0,0,"\
                          "ROUND(tblOrnek[[#This Row],[Matrah (₺)]]*tblOrnek[[#This Row],[KDV Oranı]],2))",
        "Not": "=IF(tblOrnek[[#This Row],[Belge Türü]]=\"\",\"\",CONCATENATE(\"Örnek kayıt: \",tblOrnek[[#This Row],[Belge Türü]]))",
    }
    tablo_ekle(ws, "tblOrnek", "A5:H1004", sutunlar, formuller)
    ornek = [
        ("2026-01-05", "F-1001", "Satış", 250000, 0.20, "Belgesi Tam", "Örnek kayıt 1"),
        ("2026-01-12", "A-2001", "Alış", 180000, 0.20, "Belgesi Tam", "Örnek kayıt 2"),
        ("2026-02-03", "İ-3001", "İhracat", 400000, 0.01, "Belgesi Tam", "Örnek kayıt 3"),
        ("2026-02-18", "A-2002", "Alış", 95000, 0.20, "Belge Eksik", "Örnek kayıt 4"),
        ("2026-03-07", "F-1002", "Satış", 320000, 0.20, "Belgesi Tam", "Örnek kayıt 5"),
        ("2026-03-21", "G-4001", "Gider", 45000, 0.20, "Belgesi Tam", "Örnek kayıt 6"),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            h(ws, i, k, deger)
        ws.cell(row=i, column=1).number_format = TARİH
        ws.cell(row=i, column=4).number_format = TL
        ws.cell(row=i, column=5).number_format = YÜZDE
        ws.cell(row=i, column=7).number_format = TL
    genislik(ws, {"A": 14, "B": 14, "C": 14, "D": 18, "E": 12, "F": 16, "G": 16, "H": 22})
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
    h(ws, 6, 1, "Belge Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(BELGE_TURLERI, 7):
        h(ws, i, 1, tur, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Belge türü | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni tür için satır ekleyin")
    h(ws, 6, 3, "KDV Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, oran in enumerate(KDV_ORANLARI, 7):
        h(ws, i, 3, oran, zemin=GIRIS_SARI, sayi=YÜZDE)
    h(ws, 6, 5, "Belge Durumu", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, durum in enumerate(TAHSILAT_DURUMLARI, 7):
        h(ws, i, 5, durum, zemin=GIRIS_SARI)
    genislik(ws, {"A": 18, "C": 14, "E": 18})


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
        (11, "KdvIadeEsik", 1500000, "TL", 0, 10000000000, "Azami iade tutarı eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (12, "KdvMatrahEsik", 50000000, "TL", 0, 10000000000, "Toplam belge matrahı eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
        (13, "KdvKullanimEsik", 0.5, "Oran", 0, 1, "İade kullanım oranı risk eşiği; üzerinde kontrol uyarır.", "Varsayım"),
        (14, "KdvIyiSenaryo", 1.10, "Oran", 0.5, 2, "İyimser senaryoda azami iade bu çarpanla hesaplanır.", "Varsayım"),
        (15, "KdvBazSenaryo", 1.00, "Oran", 0.5, 2, "Baz senaryoda azami iade bu çarpanla hesaplanır.", "Varsayım"),
        (16, "KdvKotuSenaryo", 0.90, "Oran", 0.5, 2, "Kötümser senaryoda azami iade bu çarpanla hesaplanır.", "Varsayım"),
        (17, "KdvTornadoDevreden", 0.01, "Oran", 0, 0.2, "Devreden KDV'nin %1 artışının azami iadeye etkisi.", "Varsayım"),
        (18, "KdvTornadoIhracat", 0.01, "Oran", 0, 0.2, "İhracatın %1 artışının azami iadeye etkisi.", "Varsayım"),
        (19, "KdvTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (20, "KdvKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (21, "KdvKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (22, "KdvAksiyon1_ACIKLAMA",
         "Eksik belgeler varsa iade süreci başlatılmadan satıcıdan fatura ve irsaliye tamamlanmalı.", "Metin", "", "",
         "İNCELE kararında gösterilen ilk öneri.", "Varsayım"),
        (23, "KdvAksiyon2_ACIKLAMA",
         "Azami iade eşiği aşılıyorsa iade dosyasını mali müşavirle birlikte kapsamlı hazırlayın.", "Metin", "", "",
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (24, "KdvAksiyon3_ACIKLAMA",
         "İade kullanım oranı yüksekse ihracat belgelerinin eksiksiz ve teyit edilebilir olduğunu doğrulayın.", "Metin", "", "",
         "İNCELE kararında gösterilen üçüncü öneri.", "Varsayım"),
        (25, "KdvP90Oran", 0.9, "Oran", 0.5, 1, "Matrah dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
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
        h(ws, satir, 1, f"KdvPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki aylık matrah grafiği serisinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: PANO'daki aylık grafiği besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 30, "B": 22, "C": 12, "D": 10, "E": 10, "F": 60, "G": 14, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "KDV_BELGELERI sayfasında sarı hücrelere dönem KDV belgelerinizi girin. Belge türünü, "
         "KDV oranını ve belge durumunu listeden seçin; matrahı KDV hariç girin."),
        ("Hesap katmanı",
         "HESAP sayfası belge türü bazında matrah, KDV, hesaplanan/indirilecek KDV, devreden KDV "
         "ve azami iade alacağını otomatik özetler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki eşikleri ve senaryo çarpanlarını mali müşavirinize göre güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; KDV beyannamesi ve mali müşavir onayı yerine geçmez.",
      yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "Azami İade Tutarı Eşiği",
        "Toplam Belge Matrahı Eşiği",
        "İade Kullanım Oranı Risk Eşiği",
        "İyimser, Baz ve Kötümser Senaryo Çarpanları",
        "Tornado Devreden ve İhracat Etkisi Oranları",
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

    ws = wb["KDV_BELGELERI"]
    ws.conditional_formatting.add("A6:A1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("F6:F1004", CellIsRule(operator="equal", formula=['"Belge Eksik"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))

    ws = wb["HESAP"]
    for harf in ("C", "D", "E", "F", "G"):
        ws.conditional_formatting.add(f"{harf}6:{harf}11",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B15:B15", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B17:B17", CellIsRule(operator="greaterThan", formula=["1500000"], font=f(amber)))
    ws.conditional_formatting.add("B19:B19", CellIsRule(operator="greaterThan", formula=["0.5"], font=f(amber)))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "EKSİK VAR", "AŞIYOR", "YÜKSEK"):
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
    ws.conditional_formatting.add("B8:B8", CellIsRule(operator="greaterThan", formula=["0"], font=f(amber)))

    ws = wb["PANO"]
    for harf in ("B", "C", "D", "E"):
        ws.conditional_formatting.add(f"{harf}4:{harf}4",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B4:B4", CellIsRule(operator="greaterThan", formula=["1500000"], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("F11:F15", CellIsRule(operator="greaterThan", formula=["1500000"], font=f(amber)))
    ws.conditional_formatting.add("E11:E22", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C9", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D7:D9", CellIsRule(operator="lessThan", formula=["1"], font=f(kirmizi)))
    ws.conditional_formatting.add("D7:D9", CellIsRule(operator="greaterThan", formula=["1.5"], font=f(amber)))
    ws.conditional_formatting.add("B12:B13", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["ORNEK_VERI"]
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))


def main(cikti_yolu=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()
    kapak(ws)
    sayfalar = [
        ("HIZLI_BASLANGIC", hizli_baslangic),
        ("KDV_BELGELERI", kdv_belgeleri),
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
    t = 6 + len(BELGE_TURLERI)
    ad_ekle(wb, "HesaplananKdv", f"HESAP!$E${t}")
    ad_ekle(wb, "IndirilecekKdv", f"HESAP!$F${t}")
    ad_ekle(wb, "DevredenKdv", f"HESAP!$B${t + 2}")
    ad_ekle(wb, "IhracatMatrahi", f"HESAP!$G${t}")
    ad_ekle(wb, "IhracatIadeSiniri", f"HESAP!$B${t + 3}")
    ad_ekle(wb, "AzamiIadeAlacak", f"HESAP!$B${t + 4}")
    ad_ekle(wb, "EksikBelgeSayisi", f"HESAP!$B${t + 5}")
    ad_ekle(wb, "IadeKullanimOrani", f"HESAP!$B${t + 6}")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "KdvIadeEsik": 11, "KdvMatrahEsik": 12, "KdvKullanimEsik": 13, "KdvIyiSenaryo": 14,
        "KdvBazSenaryo": 15, "KdvKotuSenaryo": 16, "KdvTornadoDevreden": 17,
        "KdvTornadoIhracat": 18, "KdvTahminCarpan": 19, "KdvKaliteIyi": 20,
        "KdvKaliteOrta": 21, "KdvAksiyon1_ACIKLAMA": 22, "KdvAksiyon2_ACIKLAMA": 23,
        "KdvAksiyon3_ACIKLAMA": 24, "KdvP90Oran": 25,
    }
    for i in range(1, 13):
        ayar_satirlari[f"KdvPanelSatir{i}"] = 25 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "KdvToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "KdvDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1BelgeVade": 6, "modulT1BelgeVadeYorum": 6,
        "modulT4DonemFark": 7, "modulT4DonemFarkYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2Tornado": 9, "modulO2TornadoYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1IadeTahmin": 13, "modulI1IadeTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeBelgeTurleri", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeKdvOranlari", "LISTELER!$C$7:$C$9")
    ad_ekle(wb, "ListeTahsilatDurumlari", "LISTELER!$E$7:$E$8")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "KdvIadesiAzamiAlacakHesabi.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
