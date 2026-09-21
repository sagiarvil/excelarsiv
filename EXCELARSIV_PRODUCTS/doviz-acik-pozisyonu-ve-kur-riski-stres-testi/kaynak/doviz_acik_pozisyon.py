"""
Döviz Açık Pozisyonu ve Kur Riski Stres Testi — üretim betiği.
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

URUN_AD = "Döviz Açık Pozisyonu ve Kur Riski Stres Testi"
SURUM = "1.0.0"
RENK = "1F6E8C"

DOVIZ_CINSLERI = ["USD", "EUR", "GBP", "CHF", "Diğer"]
ISLEM_TURLERI = ["Varlık", "Borç"]
VADE_DILIMLERI = ["0-30 Gün", "31-90 Gün", "91-180 Gün", "181+ Gün"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Döviz cinsi bazında varlık ve borçlarınızı işleyin; açık pozisyonunuzu ve "
                "kur şoklarında oluşacak etkiyi stres testiyle görün.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Döviz cinsi bazında varlık-borç kaydı ve net açık pozisyon hesabı",
        "Kur şoklarının şirket özkaynağına etkisini gösteren stres testi",
        "Vade yapısına göre kur riski yoğunlaşma analizi",
        "Hangi döviz cinsinin riski en çok artırdığını gösteren tornado",
        "UYGUN / İNCELE / DURDUR karar kapısı ve yönetici raporu",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Dövizli alacak-borcu olan ihracatçı ve ithalatçılar",
        "Kredi kuruluşlarından döviz kredisi kullananlar",
        "Kur hareketlerinin etkisini ölçmek isteyen finans yöneticileri",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. DOVIZ_POZISYONLARI sayfasına dövizli işlemlerinizi girin. "
                 "3. PANO ve RAPOR sayfalarından açık pozisyonu ve stres etkisini izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Döviz pozisyonlarını girin",
         "DOVIZ_POZISYONLARI sayfasındaki sarı hücrelere döviz cinsi, işlem türü (Varlık/Borç), "
         "tutar, işlem kuru ve vade bilgisini girin. Her satır bir pozisyondur."),
        ("Adım 2 — Açık pozisyonu izleyin",
         "HESAP sayfası döviz cinsi bazında varlık, borç, net pozisyon ve TL değerini hesaplar."),
        ("Adım 3 — Stres testi ve karar",
         "SENARYO_DUYARLILIK kur şoklarını simüle eder; KARAR sayfası UYGUN/İNCELE/DURDUR üretir."),
        ("Adım 4 — Eşikleri ayarlayın",
         "Açık pozisyon ve stres eşikleri AYARLAR sayfasından değiştirilir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi=KRITIK)
    hatalar = [
        "Borçları negatif tutarla girmek (borç türü pozitif tutarla işlenir)",
        "İşlem kurunu KDV dahil ile hariç karıştırmak",
        "Vade bilgisini boş bırakmak (vade yapısı analizi bozulur)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi=KRITIK)
    h(ws, 15, 1, "Bu dosya karar destek aracıdır; döviz yatırım tavsiyesi değildir.",
      yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def doviz_pozisyonlari(ws):
    sayfa_hazirla(ws, "DOVIZ_POZISYONLARI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Döviz Pozisyonları", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Her satır bir dövizli işlemdir; "
                "TL değer otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Tarih", "Döviz Cinsi", "İşlem Türü", "Tutar (Döviz)", "İşlem Kuru",
                "Vade", "TL Değer (₺)", "Kayıt Açıklaması"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "TL Değer (₺)": "=IF(tblDoviz[[#This Row],[Tutar (Döviz)]]=0,0,"\
                        "tblDoviz[[#This Row],[Tutar (Döviz)]]*tblDoviz[[#This Row],[İşlem Kuru]])",
        "Kayıt Açıklaması": "=IF(tblDoviz[[#This Row],[Döviz Cinsi]]=\"\",\"\","\
                            "CONCATENATE(\"Pozisyon: \",tblDoviz[[#This Row],[Döviz Cinsi]],\"/\",tblDoviz[[#This Row],[İşlem Türü]],\" tutar \",TEXT(tblDoviz[[#This Row],[Tutar (Döviz)]],\"#,##0.00\")))",
    }
    tablo_ekle(ws, "tblDoviz", "A5:H1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeDovizCinsleri", "B6:B1004",
              baslik="Döviz Cinsi", mesaj="Listeden bir döviz cinsi seçin.",
              hata_baslik="Geçersiz cins", hata_mesaj="Listede olmayan cins giremezsiniz.")
    dogrulama(ws, "list", "ListeIslemTurleri", "C6:C1004",
              baslik="İşlem Türü", mesaj="Varlık veya Borç seçin.",
              hata_baslik="Geçersiz tür", hata_mesaj="Sadece Varlık veya Borç kullanılır.")
    dogrulama(ws, "list", "ListeVadeDilimleri", "F6:F1004",
              baslik="Vade", mesaj="Listeden vade dilimi seçin.",
              hata_baslik="Geçersiz vade", hata_mesaj="Listede olmayan vade giremezsiniz.")
    dogrulama(ws, "date", "01.01.2024", "A6:A1004",
              baslik="Tarih", mesaj="GG.AA.YYYY biçiminde tarih girin.",
              hata_baslik="Geçersiz tarih", hata_mesaj="Tarih 01.01.2024 - 31.12.2027 aralığında olmalı.",
              isaret="between", f2="31.12.2027")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Tutar", mesaj="Pozitif döviz tutarı girin (tür Varlık/Borç olarak seçilir).",
              hata_baslik="Geçersiz tutar", hata_mesaj="0 ile 1.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="İşlem Kuru", mesaj="İşlemin yapıldığı döviz kurunu TL cinsinden girin.",
              hata_baslik="Geçersiz kur", hata_mesaj="0 ile 1.000.000 arasında olmalı.",
              isaret="between", f2="1000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=1).number_format = TARİH
        ws.cell(row=satir, column=4).number_format = "0.00"
        ws.cell(row=satir, column=5).number_format = "0.0000"
        ws.cell(row=satir, column=7).number_format = TL
    ornek_doviz = [
        ("2026-01-05", "USD", "Varlık", 50000, 34.5, "0-30 Gün", "Örnek: ihracat alacağı"),
        ("2026-01-12", "EUR", "Borç", 30000, 37.8, "31-90 Gün", "Örnek: ithalat borcu"),
        ("2026-02-03", "USD", "Borç", 20000, 35.1, "31-90 Gün", "Örnek: döviz kredisi"),
        ("2026-02-18", "EUR", "Varlık", 15000, 38.2, "0-30 Gün", "Örnek: ihracat alacağı"),
        ("2026-03-07", "GBP", "Borç", 10000, 44.0, "91-180 Gün", "Örnek: ithalat borcu"),
        ("2026-03-21", "USD", "Varlık", 25000, 36.0, "0-30 Gün", "Örnek: mevduat"),
        ("2026-04-10", "EUR", "Borç", 12000, 39.5, "91-180 Gün", "Örnek: finansal borç"),
        ("2026-05-02", "CHF", "Varlık", 8000, 39.0, "0-30 Gün", "Örnek: mevduat"),
        ("2026-05-19", "USD", "Borç", 15000, 37.0, "181+ Gün", "Örnek: uzun vadeli kredi"),
        ("2026-06-11", "EUR", "Varlık", 10000, 40.0, "0-30 Gün", "Örnek: ihracat alacağı"),
    ]
    for i, satir in enumerate(ornek_doviz, 6):
        ws.cell(row=i, column=1).value = date(*map(int, satir[0].split("-")))
        for k, deger in enumerate(satir[1:], 2):
            ws.cell(row=i, column=k).value = deger
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 14, "B": 14, "C": 12, "D": 16, "E": 12, "F": 14, "G": 16, "H": 22})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Döviz Cinsi Bazlı Açık Pozisyon Hesabı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki özet döviz cinslerine göre pozisyonlardan otomatik beslenir; "
                "elle değiştirilmez.", yazi=GRİ, kaydir=True)
    sutunlar = ["Döviz", "Varlık Tutar", "Borç Tutar", "Net Pozisyon", "Ortalama Kur",
                "TL Varlık (₺)", "TL Borç (₺)", "TL Net Pozisyon (₺)", "Kur Riski (₺)"]
    baslik_satiri(ws, 5, sutunlar)
    for r, cins in enumerate(DOVIZ_CINSLERI, 6):
        h(ws, r, 1, cins, kalin=True, yazi="333333")
        h(ws, r, 2, f'=SUMIFS(tblDoviz[Tutar (Döviz)],tblDoviz[Döviz Cinsi],$A{r},tblDoviz[İşlem Türü],"Varlık")',
          sayi="0.00", yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblDoviz[Tutar (Döviz)],tblDoviz[Döviz Cinsi],$A{r},tblDoviz[İşlem Türü],"Borç")',
          sayi="0.00", yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f"=B{r}-C{r}", sayi="0.00", yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f"=IF(COUNTIFS(tblDoviz[Döviz Cinsi],$A{r})=0,0,AVERAGEIFS(tblDoviz[İşlem Kuru],tblDoviz[Döviz Cinsi],$A{r}))",
          sayi="0.0000", yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f"=B{r}*E{r}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f"=C{r}*E{r}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 8, f"=F{r}-G{r}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 9, f"=IF(H{r}=0,0,ABS(H{r})*DovizStresEsik)", sayi=TL, yazi=DIS_REF_YEŞIL)
        for k in range(2, 10):
            yorum_ekle(ws, f"{chr(64 + k)}{r}",
                       f"Tanım: {cins} pozisyon metriği | "
                       f"Neden önemli: Açık pozisyon riskini gösterir | "
                       f"Doğru kullanım: Formüldür, değiştirilmez")
    t = 6 + len(DOVIZ_CINSLERI)
    h(ws, t, 1, "TOPLAM", kalin=True, yazi=KOYU_LACIVERT)
    for k in range(2, 9):
        h(ws, t, k, f"=SUM({chr(64+k)}6:{chr(64+k)}{t-1})", sayi=(TL if k in (6, 7, 8) else "0.00"),
          yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t, 9, f"=IF(H{t}=0,0,ABS(H{t})*DovizStresEsik)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 2, 1, "Net Açık Pozisyon (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 2, 2, f"=H{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 3, 1, "Mutlak Açık Pozisyon (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 3, 2, f"=ABS(H{t})", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 4, 1, "Açık Pozisyon Yönü", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 4, 2, f'=IF(H{t}=0,"DENGELİ",IF(H{t}>0,"FAZLA VARLIK","FAZLA BORÇ"))',
      yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 5, 1, "Stres Etkisi (Kötümser Şok)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 5, 2, f"=I{t}", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, t + 6, 1, "En Riskli Döviz Cinsi", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, t + 6, 2, f"=INDEX($A$6:$A${t-1},MATCH(MAX($I$6:$I${t-1}),$I$6:$I${t-1},0))",
      yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, t + 8, "Hesap satırları yalnızca döviz pozisyonlarından beslenir; korumalıdır.")
    genislik(ws, {"A": 18, "B": 14, "C": 14, "D": 14, "E": 12, "F": 16, "G": 16, "H": 18, "I": 18})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Vade Yapısı: 0-30 Gün Payı",
         '=IFERROR(IF(SUM(tblDoviz[Tutar (Döviz)])=0,0,SUMIFS(tblDoviz[Tutar (Döviz)],tblDoviz[Vade],"0-30 Gün")/SUM(tblDoviz[Tutar (Döviz)])),0)',
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kısa vadeli (0-30 gün) pozisyon payı ",TEXT(C6,"0.0%")," — pay yüksekse kur riski hızla realize olur")'),
        ("T2", "Ortalama İşlem Kuru",
         "=IFERROR(AVERAGE(tblDoviz[İşlem Kuru]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama işlem kuru ",TEXT(C7,"0.0000")," TL — kur eğiliminin ortalama seviyesini gösterir")'),
        ("O1", "Kur Sapması (STDEV)",
         "=IFERROR(STDEV.P(tblDoviz[İşlem Kuru]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İşlem kurlarındaki standart sapma ",TEXT(C8,"0.0000")," — sapma yüksekse kur dalgalanması belirgindir")'),
        ("O2", "Tornado: En Büyük Stres Etkisi",
         "=MAX(HESAP!I6:HESAP!I10)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kötümser şokta en yüksek etki ",TEXT(C9,"₺ #,##0")," TL — öncelik bu döviz cinsinin izlenmesinde")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C7-SENARYO_DUYARLILIK!C10",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kritik senaryo TL etkisi farkı ",TEXT(C10,"₺ #,##0")," TL — belirsizlik aralığını gösterir")'),
        ("O6", "Döviz Yoğunlaşması (HHI)",
         "=IFERROR(SUMPRODUCT((HESAP!$H$6:$H$10/SUM(ABS(HESAP!$H$6:$H$10)))^2),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Döviz yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça risk tek cinsten gelir")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=DovizKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Kur Seviyesi Tahmini",
         "=PANO!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem kur tahmini ",TEXT(C13,"0.0000")," TL — güven bandı için alt/üst sınır PANO sayfasındadır")'),
        ("I2", "Pozisyon Tutarı Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblDoviz[Tutar (Döviz)],DovizP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Pozisyon tutarlarının %90 ı ",TEXT(C14,"#,##0.00")," altında gerçekleşir — üst sınır nakit planı için kullanılabilir")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        satir = 6 + i
        h(ws, satir, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, satir, 2, ad, kaydir=True)
        h(ws, satir, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[satir].height = 30
    genislik(ws, {"A": 6, "B": 30, "C": 52, "D": 72})
    ws.sheet_view.showGridLines = True


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Giriş Kalite Kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki kontroller canlı formüllerle çalışır; sorun varsa uyarı üretir.",
      yazi=GRİ, kaydir=True)
    kontroller = [
        ("Boş döviz cinsi satırı var mı?", "=IF(COUNTBLANK(tblDoviz[Döviz Cinsi])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Boş tutar var mı?", "=IF(COUNTBLANK(tblDoviz[Tutar (Döviz)])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Negatif tutar var mı?", "=IF(COUNTIF(tblDoviz[Tutar (Döviz)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Negatif kur var mı?", "=IF(COUNTIF(tblDoviz[İşlem Kuru],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, KRITIK),
        ("Vade bilgisi eksik mi?", "=IF(COUNTBLANK(tblDoviz[Vade])=0,\"TEMİZ\",\"BOŞ VAR\")", NORMAL, KRITIK),
        ("Açık pozisyon eşiği aşılıyor mu?", "=IF(ABS(HESAP!B13)>DovizAcikEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Stres etkisi eşiği aşılıyor mu?", "=IF(HESAP!B16>DovizStresEtkiEsik,\"AŞIYOR\",\"NORMAL\")", KRITIK, NORMAL),
        ("Kur sapması risk eşiğinde mi?", "=IF(ANALITIK_MOTOR!C8>DovizKurSapmaEsik,\"YÜKSEK\",\"NORMAL\")", KRITIK, NORMAL),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        yorum_ekle(ws, f"B{i}", f"Tanım: {ad} | Neden önemli: Giriş kalitesini ölçer | "
                                f"Doğru kullanım: Formüldür, değiştirilmez | Yanlışsa: Kırmızı = düzeltin")
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Boş ve negatif girişlerin oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblDoviz[Döviz Cinsi])*COUNTA(tblDoviz[#Headers])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblDoviz[Döviz Cinsi])+COUNTA(tblDoviz[Tutar (Döviz)])+"
                 "COUNTA(tblDoviz[İşlem Kuru])+COUNTA(tblDoviz[Vade])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(DovizToplamGiris=0,0,DovizDoluGiris/DovizToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=DovizKaliteIyi,\"YÜKSEK\",IF(B19>=DovizKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 55})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Döviz Açık Pozisyon Riski", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Açık pozisyon, stres etkisi ve kur sapmasına göre üretilen karar.",
      yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Net Açık Pozisyon (TL)", yazi="333333")
    h(ws, 6, 2, "=HESAP!B13", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Stres Etkisi (Kötümser)", yazi="333333")
    h(ws, 7, 2, "=HESAP!B16", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Kur Sapması", yazi="333333")
    h(ws, 8, 2, "=ANALITIK_MOTOR!C8", sayi="0.0000", yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Pozisyon Yönü", yazi="333333")
    h(ws, 9, 2, "=HESAP!B15", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblDoviz[Döviz Cinsi])=0,"VERİ YOK",IF(COUNTIF(tblDoviz[Tutar (Döviz)],"<0")>0,"DURDUR",IF(ABS(HESAP!B13)>DovizAcikEsik,"DURDUR",IF(SUM(tblDoviz[TL Değer (₺)])>DovizHacimEsik,"İNCELE",IF(HESAP!B16>DovizStresEtkiEsik,"İNCELE","UYGUN")))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblDoviz[Tutar (Döviz)],"<0")>0,"Negatif tutar girişi tespit edildi; veriler mutlaka düzeltilmeli.",'\
                 'IF(ABS(HESAP!B13)>DovizAcikEsik,"Açık pozisyon eşiğin üzerinde; pozisyon küçültülmeli veya hedge planı yapılmalı.",'\
                 'IF(SUM(tblDoviz[TL Değer (₺)])>DovizHacimEsik,"Toplam işlem hacmi eşiğin üzerinde; pozisyon yoğunluğu incelenmeli.",'\
                 'IF(HESAP!B16>DovizStresEtkiEsik,"Stres etkisi eşiğin üzerinde; kur korumalı mevduat veya forward değerlendirilmeli.","Karar kapısı uygun: açık pozisyon ve stres etkisi hedef aralıkta."))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=DovizAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki eşiklerden türetilir; "
                     "döviz yatırım tavsiyesi değildir.")
    genislik(ws, {"A": 45, "B": 60})


def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Toplam Döviz Varlık (TL)", "=HESAP!F11", TL),
        (3, "Toplam Döviz Borç (TL)", "=HESAP!G11", TL),
        (4, "Net Açık Pozisyon (TL)", "=HESAP!H11", TL),
        (5, "Stres Etkisi (₺)", "=HESAP!I11", TL),
        (6, "Ortalama Kur", "=ANALITIK_MOTOR!C7", "0.0000"),
        (7, "Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En Riskli Döviz", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B17", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Grafik alanı verisi — döviz bazlı
    h(ws, 10, 1, "Döviz Cinsi Bazlı TL Değer", kalin=True, yazi=KOYU_LACIVERT)
    for i, cins in enumerate(DOVIZ_CINSLERI, 11):
        h(ws, i, 1, cins, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!H{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!I{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 4, "Aylık Toplam İşlem Tutarı (Döviz)", kalin=True, yazi=KOYU_LACIVERT)
    for i in range(12):
        h(ws, 11 + i, 4, f"AY{i+1}", yazi=GRİ)
        h(ws, 11 + i, 5, f"=IFERROR(INDEX(tblDoviz[Tutar (Döviz)],AYARLAR!$B${28 + i}),0)", sayi="0.00", yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Döviz Cinsi TL Değer ve Stres Etkisi"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_col=3, max_row=15), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g1.height, g1.width = 9, 16
    ws.add_chart(g1, "F6")
    g2 = PieChart()
    g2.title = "Döviz Varlık Dağılımı"
    g2.add_data(Reference(ws, min_col=2, min_row=10, max_row=15), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U6")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Döviz Cinsi Stres Etkisi"
    for i in range(11, 16):
        h(ws, i, 6, f"=HESAP!I{i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g3.add_data(Reference(ws, min_col=6, min_row=10, max_row=15), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g3.height, g3.width = 8, 13
    ws.add_chart(g3, "F19")
    g4 = LineChart()
    g4.title = "Aylık İşlem Tutarı Trendi"
    g4.add_data(Reference(ws, min_col=5, min_row=10, max_row=23), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=4, min_row=11, max_row=22))
    g4.height, g4.width = 9, 16
    ws.add_chart(g4, "U19")
    # Zaman serisi projeksiyonu
    h(ws, 24, 1, "Gelecek 1 Ay Kur Tahmini (Tahmin Aralıklı)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 24, 2, '=ROUND(IFERROR(FORECAST.LINEAR(DovizPanelSatir12,tblDoviz[İşlem Kuru],ROW(tblDoviz[İşlem Kuru])-ROW(tblDoviz[[#Headers],[İşlem Kuru]])),0),4)',
      sayi="0.0000", yazi=DIS_REF_YEŞIL)
    h(ws, 25, 1, "Son 12 Ay Ortalama Kur", yazi="333333")
    h(ws, 25, 2, "=ROUND(IFERROR(AVERAGE(tblDoviz[İşlem Kuru]),0),4)", sayi="0.0000", yazi=DIS_REF_YEŞIL)
    h(ws, 26, 1, "Son 12 Ay Kur Standart Sapma", yazi="333333")
    h(ws, 26, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblDoviz[İşlem Kuru],DovizPanelSatir1):INDEX(tblDoviz[İşlem Kuru],DovizPanelSatir12)),0),4)",
      sayi="0.0000", yazi=DIS_REF_YEŞIL)
    h(ws, 27, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 27, 2, "=B25-B26*DovizTahminCarpan", sayi="0.0000", yazi=DIS_REF_YEŞIL)
    h(ws, 28, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 28, 2, "=B25+B26*DovizTahminCarpan", sayi="0.0000", yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g5.add_data(Reference(ws, min_col=2, min_row=24, max_row=28), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=25, max_row=28))
    g5.height, g5.width = 8, 13
    ws.add_chart(g5, "F33")
    g6 = LineChart()
    g6.title = "Döviz Cinsi TL Net Pozisyon"
    g6.add_data(Reference(ws, min_col=2, min_row=10, max_row=15), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g6.height, g6.width = 8, 13
    ws.add_chart(g6, "U33")
    # Senaryo karşılaştırma verisi
    h(ws, 29, 3, "Senaryo TL Etkisi", kalin=True, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 30):
        h(ws, i, 3, f"=SENARYO_DUYARLILIK!C{i-23}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, ad, yazi=GRİ)
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Senaryo TL Etkisi Karşılaştırması"
    g7.add_data(Reference(ws, min_col=3, min_row=29, max_row=33), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=4, min_row=30, max_row=33))
    g7.height, g7.width = 8, 16
    ws.add_chart(g7, "F39")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Varlık, Borç ve Net Pozisyon Yapısı"
    g8.add_data(Reference(ws, min_col=2, min_row=3, max_row=6), titles_from_data=True)
    g8.height, g8.width = 8, 13
    ws.add_chart(g8, "U39")
    baski_hazirla(ws, "A1:M56", URUN_AD)
    alt_bant(ws, 56, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 45, "B": 22, "C": 22, "D": 22, "E": 22, "F": 22})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", "8064A2", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "İyimser, baz, kötümser ve kritik kur senaryolarında TL etkisi; ardından "
                "hangi değişkenin sonucu en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    senaryolar = [
        ("İYİMSER", "DovizIyiSenaryo"),
        ("BAZ", "DovizBazSenaryo"),
        ("KÖTÜMSER", "DovizKotuSenaryo"),
        ("KRİTİK", "DovizKritikSenaryo"),
    ]
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Kur Şok Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "TL Etkisi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 4, "Pozisyon Etki Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, (ad, carpan) in enumerate(senaryolar, 7):
        h(ws, i, 1, ad, kalin=True, yazi="333333")
        h(ws, i, 2, f"=IF(SUM(tblDoviz[Tutar (Döviz)])=0,0,{carpan})", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=ROUND(ABS(HESAP!B13)*{carpan},0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(HESAP!B13=0,0,ABS(C{i}/HESAP!B13))", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Tornado: TL Etkiye Katkı", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    tornadolar = [
        ("Kur +%1 Etkisi", "=ROUND(ABS(HESAP!B13)*DovizTornadoKur,0)", "Kur Şoku"),
        ("Pozisyon +%1 Etkisi", "=ROUND(ABS(HESAP!B13)*DovizTornadoPozisyon,0)", "Pozisyon"),
    ]
    for i, (ad, form, etiket) in enumerate(tornadolar, 13):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, "=IF(B13>=B14,\"Kur Şoku\",\"Pozisyon\")", yazi=DIS_REF_YEŞIL, kalin=True)
    alt_bant(ws, 18, "Senaryo şokları AYARLAR'daki çarpanlardan türetilir.")
    genislik(ws, {"A": 45, "B": 30, "C": 20, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "DÖVİZ AÇIK POZİSYONU VE KUR RİSKİ — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(RaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",RaporHazirlayan,\" | Sürüm: \",DosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Net Açık Pozisyon (TL)", "=HESAP!B13", TL),
        ("Mutlak Açık Pozisyon (TL)", "=HESAP!B14", TL),
        ("Pozisyon Yönü", "=HESAP!B15", None),
        ("Stres Etkisi (₺)", "=HESAP!B16", TL),
        ("Ortalama İşlem Kuru", "=ANALITIK_MOTOR!C7", "0.0000"),
        ("Kur Sapması", "=ANALITIK_MOTOR!C8", "0.0000"),
        ("En Riskli Döviz Cinsi", "=HESAP!B17", None),
        ("Döviz Yoğunlaşması (HHI)", "=ANALITIK_MOTOR!C11", "0.000"),
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
        "Net pozisyon; varlık eksi borç olarak hesaplanır; pozitif fazla varlık, negatif fazla borç demektir.",
        "Kur şokları AYARLAR'daki senaryo çarpanlarından türetilir; gerçekleşmiş veri değildir.",
        "Kur riski, pozisyonun kur şokuna duyarlılığını gösterir; türev koruma (hedge) maliyeti dahil değildir.",
        "Bu dosya döviz yatırım tavsiyesi değildir; kur riski karar destek aracıdır.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 23, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 24 + i, 1, f"=DovizAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    sutunlar = ["Tarih", "Döviz Cinsi", "İşlem Türü", "Tutar (Döviz)", "İşlem Kuru",
                "Vade", "TL Değer (₺)", "Not"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "TL Değer (₺)": "=IF(tblOrnek[[#This Row],[Tutar (Döviz)]]=0,0,"\
                        "tblOrnek[[#This Row],[Tutar (Döviz)]]*tblOrnek[[#This Row],[İşlem Kuru]])",
        "Not": "=IF(tblOrnek[[#This Row],[Döviz Cinsi]]=\"\",\"\",CONCATENATE(\"Örnek kayıt: \",tblOrnek[[#This Row],[Döviz Cinsi]]))",
    }
    tablo_ekle(ws, "tblOrnek", "A5:H1004", sutunlar, formuller)
    ornek = [
        ("2026-01-05", "USD", "Varlık", 50000, 34.5, "0-30 Gün", "Örnek kayıt 1"),
        ("2026-01-12", "EUR", "Borç", 30000, 37.8, "31-90 Gün", "Örnek kayıt 2"),
        ("2026-02-03", "USD", "Borç", 20000, 35.1, "31-90 Gün", "Örnek kayıt 3"),
        ("2026-02-18", "EUR", "Varlık", 15000, 38.2, "0-30 Gün", "Örnek kayıt 4"),
        ("2026-03-07", "GBP", "Borç", 10000, 44.0, "91-180 Gün", "Örnek kayıt 5"),
        ("2026-03-21", "USD", "Varlık", 25000, 36.0, "0-30 Gün", "Örnek kayıt 6"),
    ]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir, 1):
            h(ws, i, k, deger)
        ws.cell(row=i, column=1).number_format = TARİH
        ws.cell(row=i, column=4).number_format = "0.00"
        ws.cell(row=i, column=5).number_format = "0.0000"
        ws.cell(row=i, column=7).number_format = TL
    genislik(ws, {"A": 14, "B": 14, "C": 12, "D": 16, "E": 12, "F": 14, "G": 16, "H": 22})
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
    h(ws, 6, 1, "Döviz Cinsi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, cins in enumerate(DOVIZ_CINSLERI, 7):
        h(ws, i, 1, cins, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Döviz cinsi | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni cins için satır ekleyin")
    h(ws, 6, 3, "İşlem Türü", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, tur in enumerate(ISLEM_TURLERI, 7):
        h(ws, i, 3, tur, zemin=GIRIS_SARI)
    h(ws, 6, 5, "Vade Dilimi", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, vade in enumerate(VADE_DILIMLERI, 7):
        h(ws, i, 5, vade, zemin=GIRIS_SARI)
    genislik(ws, {"A": 18, "C": 14, "E": 16})


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
        (11, "DovizAcikEsik", 2000000, "TL", 0, 1000000000, "Açık pozisyon TL eşiği; üzerinde karar DURDUR olur.", "Varsayım"),
        (12, "DovizStresEsik", 0.2, "Oran", 0.01, 1, "Kötümser senaryo kur şoku oranı (%20).", "Varsayım"),
        (13, "DovizStresEtkiEsik", 400000, "TL", 0, 1000000000, "Stres etkisi bu eşiği aşarsa karar İNCELE olur.", "Varsayım"),
        (14, "DovizKurSapmaEsik", 0.05, "Oran", 0, 1, "İşlem kuru standart sapması risk eşiği.", "Varsayım"),
        (15, "DovizIyiSenaryo", 0.05, "Oran", 0, 1, "İyimser senaryo kur şoku (%5 lehine).", "Varsayım"),
        (16, "DovizBazSenaryo", 0.00, "Oran", -0.5, 0.5, "Baz senaryo kur şoku (değişim yok).", "Varsayım"),
        (17, "DovizKotuSenaryo", 0.10, "Oran", 0, 1, "Kötümser senaryo kur şoku (%10 aleyhine).", "Varsayım"),
        (18, "DovizKritikSenaryo", 0.30, "Oran", 0, 1, "Kritik senaryo kur şoku (%30 aleyhine).", "Varsayım"),
        (19, "DovizTornadoKur", 0.01, "Oran", 0, 0.2, "Kur şokunun %1 artışının TL etkisine katkısı.", "Varsayım"),
        (20, "DovizTornadoPozisyon", 0.01, "Oran", 0, 0.2, "Pozisyonun %1 artışının TL etkisine katkısı.", "Varsayım"),
        (21, "DovizTahminCarpan", 1.645, "Oran", 0, 3, "Tahmin bandı genişliği (standart sapma katı).", "Varsayım"),
        (22, "DovizKaliteIyi", 90, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse YÜKSEK kalite.", "Varsayım"),
        (23, "DovizKaliteOrta", 70, "Skor", 0, 100, "Kalite skoru bu eşiğin üzerindeyse ORTA kalite.", "Varsayım"),
        (24, "DovizAksiyon1_ACIKLAMA",
         "Açık pozisyon eşiği aşılıyorsa pozisyonu küçültün; döviz borcunu TL gelirle dengeleyin.", "Metin", "", "",
         "DURDUR kararında gösterilen ilk öneri.", "Varsayım"),
        (25, "DovizAksiyon2_ACIKLAMA",
         "Stres etkisi yüksekse kur korumalı mevduat veya forward sözleşmesi değerlendirin.", "Metin", "", "",
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (26, "DovizAksiyon3_ACIKLAMA",
         "Kur sapması yüksekse döviz alacaklarınızı hızlandırın ve ithalat planınızı gözden geçirin.", "Metin", "", "",
         "İNCELE kararında gösterilen üçüncü öneri.", "Varsayım"),
        (27, "DovizP90Oran", 0.9, "Oran", 0.5, 1, "Pozisyon tutarı dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (28, "DovizHacimEsik", 10000000, "TL", 0, 10000000000, "Toplam işlem hacmi (varlık+borç TL) eşiği; üzerinde karar İNCELE olur.", "Varsayım"),
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
        h(ws, satir, 1, f"DovizPanelSatir{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki aylık işlem grafiği serisinin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
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
         "DOVIZ_POZISYONLARI sayfasında sarı hücrelere dövizli işlemlerinizi girin. Döviz cinsini, "
         "işlem türünü ve vadeyi listeden seçin; tutar ve kur bilgisini girin."),
        ("Hesap katmanı",
         "HESAP sayfası döviz cinsi bazlı varlık, borç, net pozisyon, ortalama kur ve stres etkisini "
         "otomatik özetler."),
        ("Karar ve rapor",
         "KARAR sayfası UYGUN/İNCELE/DURDUR kararını üretir; RAPOR yöneticiye sunulabilir."),
        ("Ayarlar",
         "AYARLAR sayfasındaki eşikleri ve senaryo çarpanlarını finansal danışmanınıza göre güncelleyin."),
        ("Koruma",
         "Tüm sayfalar 1234 şifresiyle korunur; sarı hücreler serbesttir, formül hücreleri kilitlidir."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Bu dosya karar destek aracıdır; döviz yatırım tavsiyesi ve muhasebe kaydı niteliği taşımaz.",
      yazi=GRİ, boyut=9, kaydir=True)
    h(ws, 13, 1, "Varsayım İşaretli Parametreler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Aşağıdaki eşik ve çarpanlar AYARLAR sayfasında kaynak=Varsayım olarak işaretlidir; "
                "mevzuat zorunluluğu yoktur ve işletmeye göre güncellenmelidir.",
      yazi="333333", kaydir=True)
    ws.merge_cells(start_row=14, start_column=2, end_row=14, end_column=20)
    varsayimlar = [
        "Açık Pozisyon TL Eşiği",
        "Kötümser ve Kritik Senaryo Kur Şok Oranları",
        "Stres Etki Eşiği",
        "Kur Sapması Risk Eşiği",
        "İyimser Senaryo Kur Şoku Oranı",
        "Tornado Kur ve Pozisyon Etkisi Oranları",
        "Tahmin Çarpanı (standart sapma katı)",
        "Kalite Skoru İyi ve Orta Eşikleri",
        "P90 Yüzdelik Oranı",
        "Aksiyon Önerisi Açıklamaları",
    ]
    for j, v in enumerate(varsayimlar, 15):
        h(ws, j, 1, "• " + v, yazi="333333", boyut=9, kaydir=True)
    h(ws, 26, 1, "Bu dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. "
                "Makro ve dış bağlantı yoktur.", yazi=GRİ, boyut=9, kaydir=True)
    ws.merge_cells(start_row=26, start_column=2, end_row=26, end_column=20)
    genislik(ws, {"A": 45, "B": 60})


def kosullu_bicimlendirme(wb):
    yesil, amber, kirmizi = "1F7A4D", "B7791F", "B3261E"

    def f(renk):
        return Font(name=FONT, color=renk)

    def d(renk):
        return PatternFill("solid", start_color=renk, end_color=renk)

    ws = wb["DOVIZ_POZISYONLARI"]
    ws.conditional_formatting.add("A6:A1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("E6:E1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("G6:G1004", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))

    ws = wb["HESAP"]
    for harf in ("B", "C", "D", "F", "G", "H", "I"):
        ws.conditional_formatting.add(f"{harf}6:{harf}11",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("D6:D11", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi),
                                fill=d("FDE9E9")))
    ws.conditional_formatting.add("D6:D11", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("I6:I11", CellIsRule(operator="greaterThan", formula=["400000"], font=f(amber)))
    ws.conditional_formatting.add("I6:I11", CellIsRule(operator="greaterThan", formula=["1000000"], font=f(kirmizi)))

    ws = wb["KONTROLLER"]
    for deger in ("BOŞ VAR", "NEGATİF VAR", "AŞIYOR", "YÜKSEK"):
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
    ws.conditional_formatting.add("B7:B7", CellIsRule(operator="greaterThan", formula=["400000"], font=f(amber)))

    ws = wb["PANO"]
    for harf in ("B", "C", "D", "E"):
        ws.conditional_formatting.add(f"{harf}4:{harf}4",
                                      CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B4:B4", CellIsRule(operator="greaterThan", formula=["2000000"], font=f(kirmizi)))
    ws.conditional_formatting.add("E4:E4", CellIsRule(operator="greaterThan", formula=["400000"], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"UYGUN"'],
                                 font=f(yesil), fill=d("E2EFDA")))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"İNCELE"'], font=f(amber)))
    ws.conditional_formatting.add("B6:B6", CellIsRule(operator="equal", formula=['"DURDUR"'],
                                 font=f(kirmizi), fill=d("FDE9E9")))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B11:B15", CellIsRule(operator="greaterThan", formula=["0"], font=f(yesil)))
    ws.conditional_formatting.add("F11:F15", CellIsRule(operator="greaterThan", formula=["400000"], font=f(amber)))
    ws.conditional_formatting.add("E11:E22", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

    ws = wb["SENARYO_DUYARLILIK"]
    ws.conditional_formatting.add("C7:C10", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))
    ws.conditional_formatting.add("B7:B10", CellIsRule(operator="greaterThan", formula=["0.2"], font=f(kirmizi)))
    ws.conditional_formatting.add("B13:B14", CellIsRule(operator="lessThan", formula=["0"], font=f(kirmizi)))

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
        ("DOVIZ_POZISYONLARI", doviz_pozisyonlari),
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
    t = 6 + len(DOVIZ_CINSLERI)
    ad_ekle(wb, "NetAcikPozisyon", f"HESAP!$H${t}")
    ad_ekle(wb, "ToplamVarlikTL", f"HESAP!$F${t}")
    ad_ekle(wb, "ToplamBorcTL", f"HESAP!$G${t}")
    ad_ekle(wb, "MutlakAcikPozisyon", f"HESAP!$B${t + 3}")
    ad_ekle(wb, "AcikPozisyonYonu", f"HESAP!$B${t + 4}")
    ad_ekle(wb, "StresEtkisi", f"HESAP!$B${t + 5}")
    ad_ekle(wb, "EnRiskliDoviz", f"HESAP!$B${t + 6}")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "FirmaUnvan": 7, "RaporTarihi": 8, "RaporHazirlayan": 9, "DosyaSurumu": 10,
        "DovizAcikEsik": 11, "DovizStresEsik": 12, "DovizStresEtkiEsik": 13,
        "DovizKurSapmaEsik": 14, "DovizIyiSenaryo": 15, "DovizBazSenaryo": 16,
        "DovizKotuSenaryo": 17, "DovizKritikSenaryo": 18, "DovizTornadoKur": 19,
        "DovizTornadoPozisyon": 20, "DovizTahminCarpan": 21, "DovizKaliteIyi": 22,
        "DovizKaliteOrta": 23, "DovizAksiyon1_ACIKLAMA": 24,
        "DovizAksiyon2_ACIKLAMA": 25, "DovizAksiyon3_ACIKLAMA": 26, "DovizP90Oran": 27,
        "DovizHacimEsik": 28,
    }
    for i in range(1, 13):
        ayar_satirlari[f"DovizPanelSatir{i}"] = 28 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "DovizToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "DovizDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1VadeYapisi": 6, "modulT1VadeYapisiYorum": 6,
        "modulT2KurTrend": 7, "modulT2KurTrendYorum": 7,
        "modulO1Anomali": 8, "modulO1AnomaliYorum": 8,
        "modulO2Tornado": 9, "modulO2TornadoYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1KurTahmin": 13, "modulI1KurTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListeDovizCinsleri", "LISTELER!$A$7:$A$11")
    ad_ekle(wb, "ListeIslemTurleri", "LISTELER!$C$7:$C$8")
    ad_ekle(wb, "ListeVadeDilimleri", "LISTELER!$E$7:$E$10")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "DovizAcikPozisyonKurRiski.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
