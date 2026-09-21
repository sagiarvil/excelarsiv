"""
Pazaryeri Net Kâr & Eksik Hakediş Yakalayıcı — üretim betiği.
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

URUN_AD = "Pazaryeri Net Kâr & Eksik Hakediş Yakalayıcı"
SURUM = "1.0.0"
RENK = "2E75B6"

PAZARYERLERI = ["Trendyol", "Hepsiburada", "Amazon", "N11", "Diğer"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Pazaryerinde sattığınız ürünlerin komisyon, kargo, reklam, iade ve ek gider "
                "sonrası gerçek net kârını hesaplayın; pazaryerinden gelen tahsilatı beklenen "
                "hakedişle karşılaştırarak eksik hakedişleri otomatik yakalayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Kalem bazında satış, komisyon, kargo, reklam, iade ve ek gider sonrası net kâr ve kâr marjı hesabı",
        "Beklenen hakediş (pazaryerinin ödemesi gereken) ile gerçek hakediş farkını yakalar; eksik tahsilatı eşiklerle işaretler",
        "Pazaryeri bazında satış, kâr ve marj özeti; yoğunlaşma (HHI) ve anomali analizi",
        "Satış fiyatı senaryoları ve komisyon duyarlılığı (tornado) ile kâr bant aralığı",
        "Karar kapısı ve aksiyon önerisi ile portföy yönlendirmesi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Pazaryerinde satış yapan e-ticaret işletmeleri",
        "Ürün bazında gerçek kârlılığını izlemek isteyen operasyon yöneticileri",
        "Pazaryeri tahsilat farklarını denetleyen finans ve muhasebe ekipleri",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333")
    h(ws, 18, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. 2. KALEM_LISTESI sayfasına satış "
                 "kalemlerinizi ve gerçek hakedişleri girin. 3. PANO ve RAPOR'dan net kârı "
                 "ve eksik hakedişleri izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 21, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A21:P21")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Kalemlerinizi girin",
         "KALEM_LISTESI sayfasındaki sarı hücrelere satış kalemi bilgilerini girin: satış "
         "fiyatı, miktar, ürün maliyeti, komisyon/kargo/reklam/iade oranları, ek gider."),
        ("Adım 2 — Gerçek hakedişi girin",
         "Pazaryerinden gelen ödemeyi (hakediş) 'Gerçek Hakediş' sütununa girin; sistem "
         "beklenenle farkı hesaplar ve eksik hakedişi işaretler."),
        ("Adım 3 — Kârı izleyin",
         "HESAP ve PANO sayfaları pazaryeri bazlı satış, kâr, marj ve eksik hakediş özetini üretir."),
        ("Adım 4 — Kararı değerlendirin",
         "KARAR sayfası marj ve eksik hakediş durumuna göre GÜÇLÜ/NORMAL/İNCELE kararı üretir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Gerçek hakedişe KDV dahil tutar girmek (hakedişler KDV hariç karşılaştırılır)",
        "Komisyon oranını yüzde olarak 25 değil 0,25 girmek",
        "Miktar girilmeden satış fiyatı girmek (hesap sıfırlanır)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 16, 1, "Sık sorulan sorular", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sss = [
        "Soru: Komisyon oranını yüzde olarak nasıl girerim? Cevap: 0,25 olarak (onluk kesir). %25 yazmak hesabı 25 kat büyütür.",
        "Soru: Gerçek hakediş KDV dahil mi? Cevap: Hakediş mutabakatı KDV hariç yapılır; KDV dahil tutar girmeyin.",
        "Soru: Beklenen hakedişten hangi tutarlar düşülür? Cevap: Komisyon, kargo/operasyon ve reklam tutarları düşülür; iade maliyeti kâr/marjda etkilidir.",
        "Soru: Hangi pazaryeri listesinde yoksa ne yapmalıyım? Cevap: LISTELER sayfasındaki sarı hücreye yeni pazaryeri adını ekleyin.",
    ]
    for i, m in enumerate(sss, 17):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
        ws.row_dimensions[i].height = 32
    h(ws, 22, 1, "Bu dosya karar destek aracıdır; vergi ve muhasebe hesabı mali müşavir görüşü "
                 "yerine geçmez.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def kalem_listesi(ws):
    sayfa_hazirla(ws, "KALEM_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Satış Kalemleri ve Hakediş Takibi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Satış, komisyon, kargo, reklam, iade, "
                "toplam maliyet, net kâr, marj ve hakediş farkı otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Kalem Kodu", "Ürün Adı", "Birim Satış Fiyatı (TL)", "Miktar (Adet)",
                "Birim Ürün Maliyeti (TL)", "Komisyon Oranı", "Kargo/Operasyon Oranı",
                "Reklam Oranı", "İade Oranı", "Birim Ek Gider (TL)",
                "Toplam Satış (TL)", "Komisyon Tutarı (TL)", "Kargo Tutarı (TL)",
                "Reklam Tutarı (TL)", "İade Maliyeti (TL)", "Toplam Maliyet (TL)",
                "NET KÂR (TL)", "KÂR MARJI", "Beklenen Hakediş (TL)", "Gerçek Hakediş (TL)",
                "Hakediş Farkı (TL)", "Fark Oranı", "Durum", "Açıklama", "Pazaryeri"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Toplam Satış (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblKalem[[#This Row],[Miktar (Adet)]])',
        "Komisyon Tutarı (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Komisyon Oranı]]*tblKalem[[#This Row],[Toplam Satış (TL)]])',
        "Kargo Tutarı (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Kargo/Operasyon Oranı]]*tblKalem[[#This Row],[Toplam Satış (TL)]])',
        "Reklam Tutarı (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Reklam Oranı]]*tblKalem[[#This Row],[Toplam Satış (TL)]])',
        "İade Maliyeti (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[İade Oranı]]*tblKalem[[#This Row],[Toplam Satış (TL)]]*(1-tblKalem[[#This Row],[Komisyon Oranı]]))',
        "Toplam Maliyet (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Birim Ürün Maliyeti (TL)]]*tblKalem[[#This Row],[Miktar (Adet)]]+tblKalem[[#This Row],[Komisyon Tutarı (TL)]]+tblKalem[[#This Row],[Kargo Tutarı (TL)]]+tblKalem[[#This Row],[Reklam Tutarı (TL)]]+tblKalem[[#This Row],[İade Maliyeti (TL)]]+tblKalem[[#This Row],[Birim Ek Gider (TL)]]*tblKalem[[#This Row],[Miktar (Adet)]])',
        "NET KÂR (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Toplam Satış (TL)]]-tblKalem[[#This Row],[Toplam Maliyet (TL)]])',
        "KÂR MARJI": '=IF(tblKalem[[#This Row],[Toplam Satış (TL)]]=0,0,tblKalem[[#This Row],[NET KÂR (TL)]]/tblKalem[[#This Row],[Toplam Satış (TL)]])',
        "Beklenen Hakediş (TL)": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="",0,tblKalem[[#This Row],[Toplam Satış (TL)]]-tblKalem[[#This Row],[Komisyon Tutarı (TL)]]-tblKalem[[#This Row],[Kargo Tutarı (TL)]]-tblKalem[[#This Row],[Reklam Tutarı (TL)]])',
        "Hakediş Farkı (TL)": '=IF(tblKalem[[#This Row],[Gerçek Hakediş (TL)]]="",0,MAX(0,tblKalem[[#This Row],[Beklenen Hakediş (TL)]]-tblKalem[[#This Row],[Gerçek Hakediş (TL)]]))',
        "Fark Oranı": '=IF(tblKalem[[#This Row],[Beklenen Hakediş (TL)]]=0,0,tblKalem[[#This Row],[Hakediş Farkı (TL)]]/tblKalem[[#This Row],[Beklenen Hakediş (TL)]])',
        "Durum": '=IF(tblKalem[[#This Row],[Miktar (Adet)]]="","",IF(tblKalem[[#This Row],[Gerçek Hakediş (TL)]]="","BEKLİYOR",IF(OR(tblKalem[[#This Row],[Hakediş Farkı (TL)]]=0,tblKalem[[#This Row],[Beklenen Hakediş (TL)]]=0),"DOĞRU",IF(AND(tblKalem[[#This Row],[Fark Oranı]]>PazaryeriEksikHakedisEsik,tblKalem[[#This Row],[Hakediş Farkı (TL)]]>PazaryeriEksikHakedisTutarEsik),"EKSİK HAKEDİŞ","DOĞRU"))))',
        "Açıklama": '=IF(tblKalem[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblKalem[[#This Row],[Ürün Adı]]," | Kâr: ",TEXT(tblKalem[[#This Row],[NET KÂR (TL)]],"₺ #,##0")," | Durum: ",tblKalem[[#This Row],[Durum]]))',
    }
    tablo_ekle(ws, "tblKalem", "A5:Y1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListePazaryerleri", "Y6:Y1004",
              baslik="Pazaryeri", mesaj="Listeden pazaryerini seçin.",
              hata_baslik="Geçersiz pazaryeri", hata_mesaj="Listede olmayan pazaryeri giremezsiniz.")
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Birim Satış Fiyatı", mesaj="Satış fiyatını TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000")
    dogrulama(ws, "whole", "1", "D6:D1004",
              baslik="Miktar", mesaj="Satılan adet miktarını girin.",
              hata_baslik="Geçersiz miktar", hata_mesaj="1 veya daha büyük tamsayı olmalı.")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Birim Ürün Maliyeti", mesaj="Birim ürün maliyetini TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000")
    for kolon in ("F", "G", "H", "I"):
        dogrulama(ws, "decimal", "0", f"{kolon}6:{kolon}1004",
                  baslik="Oran", mesaj="Oranı 0-1 arasında girin (örn. %15 için 0,15).",
                  hata_baslik="Geçersiz oran", hata_mesaj="0 ile 1 arasında olmalı.",
                  isaret="between", f2="1")
    dogrulama(ws, "decimal", "0", "J6:J1004",
              baslik="Birim Ek Gider", mesaj="Paketleme/işçilik gibi birim ek gideri TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000 arasında olmalı.",
              isaret="between", f2="1000000")
    dogrulama(ws, "decimal", "0", "T6:T1004",
              baslik="Gerçek Hakediş", mesaj="Pazaryerinden gelen gerçek ödemeyi TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000.000 arasında olmalı.",
              isaret="between", f2="1000000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = TL
        ws.cell(row=satir, column=4).number_format = CATI
        ws.cell(row=satir, column=5).number_format = TL
        for kolon in (6, 7, 8, 9):
            ws.cell(row=satir, column=kolon).number_format = YÜZDE
        ws.cell(row=satir, column=10).number_format = TL
        ws.cell(row=satir, column=11).number_format = TL
        ws.cell(row=satir, column=12).number_format = TL
        ws.cell(row=satir, column=13).number_format = TL
        ws.cell(row=satir, column=14).number_format = TL
        ws.cell(row=satir, column=15).number_format = TL
        ws.cell(row=satir, column=16).number_format = TL
        ws.cell(row=satir, column=17).number_format = TL
        ws.cell(row=satir, column=18).number_format = YÜZDE
        ws.cell(row=satir, column=19).number_format = TL
        ws.cell(row=satir, column=20).number_format = TL
        ws.cell(row=satir, column=21).number_format = TL
        ws.cell(row=satir, column=22).number_format = YÜZDE
    ornek = [
        ("P-1001", "Kulaklık Bluetooth", "Trendyol", 899, 25, 420, 0.18, 0.04, 0.06, 0.05, 12,
         "", "", "", "", "", "", "", "", "", 17400, "", "", "", ""),
        ("P-1002", "Powerbank 20.000mAh", "Hepsiburada", 549, 60, 260, 0.22, 0.05, 0.05, 0.04, 9,
         "", "", "", "", "", "", "", "", "", 22400, "", "", "", ""),
        ("P-1003", "Akıllı Saat", "Trendyol", 1499, 18, 820, 0.20, 0.04, 0.08, 0.06, 18,
         "", "", "", "", "", "", "", "", "", 19600, "", "", "", ""),
        ("P-1004", "Şarj Aleti 65W", "Amazon", 349, 120, 160, 0.24, 0.06, 0.07, 0.03, 6,
         "", "", "", "", "", "", "", "", "", 26000, "", "", "", ""),
        ("P-1005", "Kablo USB-C", "N11", 129, 300, 55, 0.15, 0.08, 0.04, 0.06, 3,
         "", "", "", "", "", "", "", "", "", 28300, "", "", "", ""),
        ("P-1006", "Mouse Kablosuz", "Amazon", 279, 80, 130, 0.25, 0.06, 0.09, 0.05, 5,
         "", "", "", "", "", "", "", "", "", 13400, "", "", "", ""),
    ]
    giris_kolonlari = [1, 2, 25, 3, 4, 5, 6, 7, 8, 9, 10]
    for i, satir in enumerate(ornek, 6):
        for k in range(11):
            ws.cell(row=i, column=giris_kolonlari[k]).value = satir[k]
        # gerçek hakediş: iki kalemde eksik tahsilat simüle edilir
        if i in (6, 9):
            ws.cell(row=i, column=20).value = None
        elif i == 7:
            ws.cell(row=i, column=20).value = 16500
        elif i == 8:
            ws.cell(row=i, column=20).value = 19600
        elif i == 10:
            ws.cell(row=i, column=20).value = 28300
        elif i == 11:
            ws.cell(row=i, column=20).value = 12750
        ws.cell(row=i, column=20).number_format = TL
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 20, 25])
    sabitle(ws, "A6")
    genislik(ws, {"A": 11, "B": 24, "C": 15, "D": 11, "E": 16, "F": 12, "G": 16,
                  "H": 11, "I": 10, "J": 14, "K": 16, "L": 15, "M": 15, "N": 13, "O": 14,
                  "P": 16, "Q": 16, "R": 10, "S": 17, "T": 17, "U": 15, "V": 10, "W": 16,
                  "X": 40, "Y": 15})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")
def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Pazaryeri Bazlı Özet ve Net Kâr Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Özet kalem listesinden otomatik beslenir; pazaryeri bazında satış, maliyet, "
                "net kâr, marj ve hakediş farkı canlı üretilir.", yazi=GRİ, kaydir=True)
    sutunlar = ["Pazaryeri", "Adet", "Toplam Satış (TL)", "Komisyon (TL)", "Kargo (TL)",
                "Reklam (TL)", "İade (TL)", "Toplam Maliyet (TL)", "NET KÂR (TL)",
                "KÂR MARJI", "Beklenen Hakediş (TL)", "Hakediş Farkı (TL)"]
    baslik_satiri(ws, 5, sutunlar)
    for r, paz in enumerate(PAZARYERLERI, 6):
        h(ws, r, 1, paz, kalin=True, yazi="333333")
        h(ws, r, 2, f'=IFERROR(COUNTIFS(tblKalem[Pazaryeri],$A{r},tblKalem[Miktar (Adet)],"<>0"),0)',
          sayi=CATI, yazi=DIS_REF_YEŞIL)
        h(ws, r, 3, f'=SUMIFS(tblKalem[Toplam Satış (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 4, f'=SUMIFS(tblKalem[Komisyon Tutarı (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 5, f'=SUMIFS(tblKalem[Kargo Tutarı (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 6, f'=SUMIFS(tblKalem[Reklam Tutarı (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 7, f'=SUMIFS(tblKalem[İade Maliyeti (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 8, f'=SUMIFS(tblKalem[Toplam Maliyet (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 9, f'=SUMIFS(tblKalem[NET KÂR (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 10, f'=IFERROR(C{r}/B{r},0)', sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, r, 11, f'=SUMIFS(tblKalem[Beklenen Hakediş (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, r, 12, f'=SUMIFS(tblKalem[Hakediş Farkı (TL)],tblKalem[Pazaryeri],$A{r})', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "TOPLAM BLOK", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    toplamlar = [
        (14, "Toplam Satış (TL)", "=SUM(C6:C10)", TL),
        (15, "Toplam Komisyon (TL)", "=SUM(D6:D10)", TL),
        (16, "Toplam Kargo/Operasyon (TL)", "=SUM(E6:E10)", TL),
        (17, "Toplam Reklam (TL)", "=SUM(F6:F10)", TL),
        (18, "Toplam İade Maliyeti (TL)", "=SUM(G6:G10)", TL),
        (19, "Toplam Maliyet (TL)", "=SUM(H6:H10)", TL),
        (20, "TOPLAM NET KÂR (TL)", "=SUM(I6:I10)", TL),
        (21, "Genel Kâr Marjı", "=IFERROR(B20/B14,0)", YÜZDE),
        (22, "Ortalama Kesinti Oranı", "=IFERROR((B15+B16+B17+B18)/B14,0)", YÜZDE),
        (23, "Toplam Beklenen Hakediş (TL)", "=SUM(K6:K10)", TL),
        (24, "TOPLAM HAKEDİŞ FARKI (TL)", "=SUM(L6:L10)", TL),
        (25, "Eksik Hakediş Kalem Sayısı", '=COUNTIF(tblKalem[Durum],"EKSİK HAKEDİŞ")', CATI),
        (26, "Ortalama Kâr Marjı", "=IFERROR(AVERAGE(tblKalem[KÂR MARJI]),0)", YÜZDE),
        (27, "En İyi Pazaryeri", "=IFERROR(INDEX(A6:A10,MATCH(MAX(J6:J10),J6:J10,0)),\"\")", None),
        (28, "En İyi Pazaryeri Marjı", "=IFERROR(MAX(J6:J10),0)", YÜZDE),
    ]
    for satir, ad, form, sayi in toplamlar:
        kalin = satir in (20, 24)
        h(ws, satir, 1, ad, yazi="333333", kalin=kalin)
        h(ws, satir, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=kalin)
    # Panel verisi (grafikler için)
    h(ws, 30, 1, "Kalem Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 30, 2, "NET KÂR (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 30, 4, "Hakediş Farkı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 31 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 31 + i, 2, f"=IFERROR(INDEX(tblKalem[NET KÂR (TL)],PazaryeriPanelSatir_{i+1}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 31 + i, 4, f"=IFERROR(INDEX(tblKalem[Hakediş Farkı (TL)],PazaryeriPanelSatir_{i+1}),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
    # Projeksiyon
    h(ws, 44, 1, "Ortalama NET KÂR (TL)", yazi="333333")
    h(ws, 44, 2, "=ROUND(IFERROR(AVERAGE(INDEX(tblKalem[NET KÂR (TL)],PazaryeriPanelSatir_1):INDEX(tblKalem[NET KÂR (TL)],PazaryeriPanelSatir_12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 45, 1, "Gelecek Dönem NET KÂR Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 45, 2, '=ROUND(IFERROR(FORECAST.LINEAR(PazaryeriPanelSatir_12,tblKalem[NET KÂR (TL)],ROW(tblKalem[NET KÂR (TL)])-ROW(tblKalem[[#Headers],[NET KÂR (TL)]])),0),0)',
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 46, 1, "NET KÂR Std Sapması (TL)", yazi="333333")
    h(ws, 46, 2, "=ROUND(IFERROR(STDEV.P(INDEX(tblKalem[NET KÂR (TL)],PazaryeriPanelSatir_1):INDEX(tblKalem[NET KÂR (TL)],PazaryeriPanelSatir_12)),0),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 47, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 47, 2, "=B44-B46*PazaryeriTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 48, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 48, 2, "=B44+B46*PazaryeriTahminCarpan", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 50, "Hesap satırları yalnızca kalem listesinden ve AYARLAR'dan beslenir; korumalıdır.")
    genislik(ws, {"A": 28, "B": 9, "C": 16, "D": 14, "E": 12, "F": 12, "G": 12, "H": 16,
                  "I": 15, "J": 10, "K": 18, "L": 16})


def analitik_motor(ws):
    sayfa_hazirla(ws, "ANALITIK_MOTOR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "ANALİTİK MOTOR — DERİNLİK KATMANI", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    ws.merge_cells("A3:D3")
    baslik_satiri(ws, 5, ["Kod", "Gösterge", "Değer", "Makine Yorumu"])
    moduller = [
        ("T1", "Ortalama Kâr Marjı",
         "=IFERROR(AVERAGE(tblKalem[KÂR MARJI]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama kâr marjı ",TEXT(C6,"0.0%")," — portföyün genel kârlılık düzeyini gösterir")'),
        ("T2", "Ortalama Komisyon Oranı",
         "=IFERROR(AVERAGE(tblKalem[Komisyon Oranı]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama komisyon oranı ",TEXT(C7,"0.0%")," — pazaryeri maliyet yükünü gösterir")'),
        ("O1", "Net Kâr Sapması (MAD)",
         "=IFERROR(AVEDEV(tblKalem[NET KÂR (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Net kâr ortalama mutlak sapması ",TEXT(C8,"₺ #,##0")," TL — kalemler arası kâr heterojenliğini ölçer")'),
        ("O2", "En Yüksek Net Kâr",
         "=IFERROR(MAX(tblKalem[NET KÂR (TL)]),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek net kâr ",TEXT(C9,"₺ #,##0")," TL — bu kalem portföyün kâr lideridir")'),
        ("O3", "Senaryo Bant Genişliği",
         "=SENARYO_DUYARLILIK!C11",
         '=_xlfn.TEXTJOIN("; ",TRUE,"İyimser ile kritik senaryo net kâr farkı ",TEXT(C10,"₺ #,##0")," TL — satış fiyatı belirsizliğinin etkisini gösterir")'),
        ("O6", "Pazaryeri Yoğunlaşması (HHI)",
         "=IFERROR(IF(SUM(HESAP!$C$6:$C$10)=0,0,SUMPRODUCT((HESAP!$C$6:$C$10/SUM(HESAP!$C$6:$C$10))^2)),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Pazaryeri yoğunlaşma endeksi ",TEXT(C11,"0.000")," — 1 e yaklaştıkça satış tek pazaryerinde toplanır")'),
        ("O8", "Veri Kalite Skoru",
         "=KONTROLLER!B19",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C12,"0")," puan — ",IF(C12>=PazaryeriKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))'),
        ("I1", "Gelecek Dönem NET KÂR Tahmini",
         "=HESAP!B45",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem net kâr tahmini ",TEXT(C13,"₺ #,##0")," TL — güven bandı için alt/üst sınır HESAP sayfasındadır")'),
        ("I2", "Satış Fiyatı Yüzdelik P90",
         "=IFERROR(_xlfn.PERCENTILE.INC(tblKalem[Birim Satış Fiyatı (TL)],PazaryeriP90Oran),0)",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Kalemlerin %90 ı ",TEXT(C14,"₺ #,##0")," TL altında satış fiyatına sahiptir — fiyat dağılımı üst sınırını gösterir")'),
        ("I8", "Toplam Eksik Hakediş Etkisi",
         "=HESAP!B24",
         '=_xlfn.TEXTJOIN("; ",TRUE,"Tespit edilen toplam eksik hakediş ",TEXT(C15,"₺ #,##0")," TL — pazaryerinden talep edilmesi gereken tutar")'),
    ]
    for i, (kod, ad, form, yorum) in enumerate(moduller):
        satir = 6 + i
        h(ws, satir, 1, kod, kalin=True, hiza="center", yazi="B08948")
        h(ws, satir, 2, ad, kaydir=True)
        h(ws, satir, 3, form, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, yorum, yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[satir].height = 30
    genislik(ws, {"A": 6, "B": 34, "C": 60, "D": 72})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", "C00000", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Giriş Kalite Kontrolleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Aşağıdaki kontroller canlı formüllerle çalışır; sorun varsa uyarı üretir.",
      yazi=GRİ, kaydir=True)
    kontroller = [
        ("Negatif satış fiyatı var mı?", "=IF(COUNTIF(tblKalem[Birim Satış Fiyatı (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Negatif ürün maliyeti var mı?", "=IF(COUNTIF(tblKalem[Birim Ürün Maliyeti (TL)],\"<0\")=0,\"TEMİZ\",\"NEGATİF VAR\")", NORMAL, "B3261E"),
        ("Komisyon oranı 0,50 üzerinde mi?", "=IF(COUNTIF(tblKalem[Komisyon Oranı],\">0.5\")=0,\"TEMİZ\",\"ORAN >0,5 VAR\")", NORMAL, "B3261E"),
        ("İade oranı 0,50 üzerinde mi?", "=IF(COUNTIF(tblKalem[İade Oranı],\">0.5\")=0,\"TEMİZ\",\"ORAN >0,5 VAR\")", NORMAL, "B3261E"),
        ("Toplam hakediş farkı eşiği aşıyor mu?", "=IF(HESAP!B24>PazaryeriToplamFarkEsik,\"AŞIYOR\",\"NORMAL\")", "B3261E", NORMAL),
        ("Eksik hakediş kalemi var mı?", "=IF(COUNTIF(tblKalem[Durum],\"EKSİK HAKEDİŞ\")>0,\"VAR\",\"YOK\")", "B3261E", NORMAL),
        ("Genel marj kritik altında mı?", "=IF(HESAP!B21<PazaryeriNetKarMarjiKritik,\"KRİTİK\",\"NORMAL\")", "B3261E", NORMAL),
        ("Kalem sayısı yeterli mi?", "=IF(COUNTA(tblKalem[Ürün Adı])>=3,\"YETERLİ\",\"AZ\")", NORMAL, "B3261E"),
    ]
    for i, (ad, form, iyi, kotu) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Veri Kalitesi Skoru (0-100)", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 16, 1, "Temel giriş kolonlarının dolu olma oranına göre üretilen kalite skoru.", yazi=GRİ, kaydir=True)
    h(ws, 17, 1, "Toplam Giriş Hücresi", yazi="333333")
    h(ws, 17, 2, "=COUNTA(tblKalem[Ürün Adı])*PazaryeriTemelKolonSayisi", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Dolu Giriş Hücresi", yazi="333333")
    h(ws, 18, 2, "=COUNTA(tblKalem[Ürün Adı])+COUNTA(tblKalem[Birim Satış Fiyatı (TL)])+"
                 "COUNTA(tblKalem[Miktar (Adet)])+COUNTA(tblKalem[Birim Ürün Maliyeti (TL)])+"
                 "COUNTA(tblKalem[Komisyon Oranı])+COUNTA(tblKalem[Birim Ek Gider (TL)])+"
                 "COUNTA(tblKalem[Gerçek Hakediş (TL)])",
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kalite Skoru", yazi="333333", kalin=True)
    h(ws, 19, 2, "=ROUND(IF(PazaryeriToplamGiris=0,0,PazaryeriDoluGiris/PazaryeriToplamGiris*100),0)",
      sayi=CATI, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kalite Seviyesi", yazi="333333")
    h(ws, 20, 2, "=IF(B19>=PazaryeriKaliteIyi,\"YÜKSEK\",IF(B19>=PazaryeriKaliteOrta,\"ORTA\",\"DÜŞÜK\"))",
      yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 22, "Kontroller yalnızca giriş verisini denetler; değerleri değiştirmez.")
    genislik(ws, {"A": 45, "B": 60})


def karar_motoru(ws):
    sayfa_hazirla(ws, "KARAR", "B3261E", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Karar Kapısı — Portföy Yönlendirme", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Genel kâr marjı, eksik hakediş durumu ve veri kalitesine göre üretilen "
                "portföy kararı.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Toplam Net Kâr (TL)", yazi="333333")
    h(ws, 6, 2, "=HESAP!B20", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "Genel Kâr Marjı", yazi="333333")
    h(ws, 7, 2, "=HESAP!B21", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "Toplam Hakediş Farkı (TL)", yazi="333333")
    h(ws, 8, 2, "=HESAP!B24", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Eksik Hakediş Kalem Sayısı", yazi="333333")
    h(ws, 9, 2, "=HESAP!B25", sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblKalem[Ürün Adı])=0,"VERİ YOK",'\
                 'IF(OR(COUNTIF(tblKalem[Birim Satış Fiyatı (TL)],"<0")>0,COUNTIF(tblKalem[Birim Ürün Maliyeti (TL)],"<0")>0),"DURDUR",'\
                 'IF(HESAP!B14<=0,"VERİ YOK",'\
                 'IF(AND(COUNTIF(tblKalem[Durum],"EKSİK HAKEDİŞ")>0,HESAP!B24>0),"İNCELE",'\
                 'IF(HESAP!B21>=PazaryeriNetKarMarjiHedef,"GÜÇLÜ",'\
                 'IF(HESAP!B21>=PazaryeriNetKarMarjiKritik,"NORMAL","İNCELE"))))))',
      boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kalin=True)
    h(ws, 12, 2, '=IF(COUNTIF(tblKalem[Birim Satış Fiyatı (TL)],"<0")>0,"Negatif satış fiyatı tespit edildi; veriler düzeltilmeli.",'\
                 'IF(COUNTIF(tblKalem[Birim Ürün Maliyeti (TL)],"<0")>0,"Negatif ürün maliyeti tespit edildi; veriler düzeltilmeli.",'\
                 'IF(HESAP!B14<=0,"Satış kaydı bulunmuyor; portföy kararı üretilemiyor.",'\
                 'IF(AND(COUNTIF(tblKalem[Durum],"EKSİK HAKEDİŞ")>0,HESAP!B24>0),"Eksik hakediş tespit edildi; pazaryerinden fark tutarı talep edilmelidir.",'\
                 'IF(HESAP!B21>=PazaryeriNetKarMarjiHedef,"Genel marj hedefin üzerinde; portföy güçlü, büyütmeye devam edin.",'\
                 'IF(HESAP!B21>=PazaryeriNetKarMarjiKritik,"Genel marj sınırın üzerinde; portföy normal, maliyet kalemlerini izleyin.",'\
                 '"Genel marj kritik altında; fiyat veya komisyon yapısı yeniden değerlendirilmeli.")))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.merge_cells("B12:J12")
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 15 + i, 1, f"Aksiyon {i+1}", yazi="333333")
        h(ws, 15 + i, 2, f"=PazaryeriAksiyon{i+1}_ACIKLAMA", yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.merge_cells(start_row=15 + i, start_column=2, end_row=15 + i, end_column=10)
    alt_bant(ws, 19, "Karar yalnızca giriş verisi ve AYARLAR'daki oran/eşiklerden türetilir.")
    genislik(ws, {"A": 45, "B": 60})
def pano(ws):
    sayfa_hazirla(ws, "PANO", "0F2742", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetim Paneli", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (2, "Toplam Satış (TL)", "=HESAP!B14", TL),
        (3, "Toplam NET KÂR (TL)", "=HESAP!B20", TL),
        (4, "Genel Kâr Marjı", "=HESAP!B21", YÜZDE),
        (5, "Toplam Beklenen Hakediş (TL)", "=HESAP!B23", TL),
        (6, "Toplam Hakediş Farkı (TL)", "=HESAP!B24", TL),
        (7, "Eksik Kalem Sayısı", "=HESAP!B25", CATI),
        (8, "Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En İyi Pazaryeri", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B27", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "En İyi Pazaryeri Marjı", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[6].height = 30
    h(ws, 6, 8, "=HESAP!B28", sayi=YÜZDE, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    # Pazaryeri bazlı grafik verisi
    h(ws, 10, 1, "Pazaryeri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Toplam Satış (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "NET KÂR (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "KÂR MARJI", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Beklenen Hakediş (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, paz in enumerate(PAZARYERLERI, 11):
        h(ws, i, 1, f"{paz} Satış", yazi=GRİ)
        h(ws, i, 2, f"=HESAP!C{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!I{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!J{i-5}", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=HESAP!K{i-5}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Pazaryeri Bazlı Toplam Satış"
    g1.add_data(Reference(ws, min_col=2, min_row=10, max_row=15), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g1.height, g1.width = 9, 15
    ws.add_chart(g1, "F8")
    g2 = PieChart()
    g2.title = "Pazaryeri Bazlı Net Kâr Payı"
    g2.add_data(Reference(ws, min_col=3, min_row=10, max_row=15), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g3 = BarChart()
    g3.type = "col"
    g3.title = "Pazaryeri Bazlı Kâr Marjı"
    g3.add_data(Reference(ws, min_col=4, min_row=10, max_row=15), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g3.height, g3.width = 9, 15
    ws.add_chart(g3, "F21")
    g9 = BarChart()
    g9.type = "col"
    g9.title = "Pazaryeri Bazlı Beklenen Hakediş"
    g9.add_data(Reference(ws, min_col=5, min_row=10, max_row=15), titles_from_data=True)
    g9.set_categories(Reference(ws, min_col=1, min_row=11, max_row=15))
    g9.height, g9.width = 9, 12
    ws.add_chart(g9, "U21")
    # Kalem bazlı panel
    h(ws, 20, 7, "Kalem Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 20, 8, "NET KÂR (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 20, 9, "Hakediş Farkı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(12):
        h(ws, 21 + i, 7, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 21 + i, 8, f"=HESAP!B{31 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 21 + i, 9, f"=HESAP!D{31 + i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Kalem Bazlı NET KÂR Trendi"
    g4.add_data(Reference(ws, min_col=8, min_row=20, max_row=32), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=7, min_row=21, max_row=32))
    g4.height, g4.width = 9, 15
    ws.add_chart(g4, "F34")
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Kalem Bazlı Hakediş Farkı"
    g5.add_data(Reference(ws, min_col=9, min_row=20, max_row=32), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=7, min_row=21, max_row=32))
    g5.height, g5.width = 9, 12
    ws.add_chart(g5, "U34")
    # Durum dağılımı
    h(ws, 34, 1, "Durum", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 34, 2, "Adet", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, durum in enumerate(["EKSİK HAKEDİŞ", "DOĞRU", "BEKLİYOR"], 35):
        h(ws, i, 1, durum, yazi=GRİ)
        h(ws, i, 2, f'=IFERROR(COUNTIF(tblKalem[Durum],"{durum}"),0)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    g7 = PieChart()
    g7.title = "Hakediş Durum Dağılımı"
    g7.add_data(Reference(ws, min_col=2, min_row=34, max_row=38), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=35, max_row=38))
    g7.height, g7.width = 9, 12
    ws.add_chart(g7, "F47")
    # Projeksiyon
    h(ws, 40, 1, "Ortalama NET KÂR", yazi="333333")
    h(ws, 40, 2, "=HESAP!B44", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 41, 1, "Gelecek Dönem NET KÂR Tahmini", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 41, 2, "=HESAP!B45", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 42, 1, "NET KÂR Std Sapması", yazi="333333")
    h(ws, 42, 2, "=HESAP!B46", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 43, 1, "Alt Tahmin Sınırı", yazi="333333")
    h(ws, 43, 2, "=HESAP!B47", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 44, 1, "Üst Tahmin Sınırı", yazi="333333")
    h(ws, 44, 2, "=HESAP!B48", sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = BarChart()
    g6.type = "col"
    g6.title = "Projeksiyon: Ortalama, Alt ve Üst Sınır"
    g6.add_data(Reference(ws, min_col=2, min_row=40, max_row=44), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=1, min_row=41, max_row=44))
    g6.height, g6.width = 9, 15
    ws.add_chart(g6, "U47")
    # Senaryo verisi
    h(ws, 46, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 46, 6, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 47):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-40}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g8 = LineChart()
    g8.title = "Senaryo Net Kârları"
    g8.add_data(Reference(ws, min_col=6, min_row=46, max_row=51), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=5, min_row=47, max_row=51))
    g8.height, g8.width = 9, 12
    ws.add_chart(g8, "F54")
    baski_hazirla(ws, "A1:M52", URUN_AD)
    alt_bant(ws, 52, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 40, "B": 20, "C": 20, "D": 18, "E": 20, "F": 20, "G": 20, "H": 20, "I": 18})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Farklı satış fiyatı çarpanı senaryolarında net kâr bant aralığı; ardından "
                "hangi değişkenin net kârı en çok etkilediğini gösteren tornado analizi.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Satış Fiyatı Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 6, 3, "Net Kâr (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
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
        h(ws, i, 3, f"=IFERROR(PazaryeriToplamNetKar+PazaryeriToplamSatis*B{i}*(1-PazaryeriOrtKesintiOrani),0)",
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=IF(C{i}>=PazaryeriToplamNetKar,\"AVANTAJLI\",\"DEĞERLENDİR\")", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "Senaryo Bant Genişliği (İyi − Kritik)", yazi="333333", kalin=True)
    h(ws, 11, 3, "=C7-C10", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 13, 1, "Tornado: Net Kâra Etki", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 14, 1, "Komisyon Oranı +1 puan Etkisi", yazi="333333")
    h(ws, 14, 2, "=IFERROR(ABS(PazaryeriToplamSatis*PazaryeriTornadoKomisyonOran),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Satış Fiyatı −%5 Etkisi", yazi="333333")
    h(ws, 15, 2, "=IFERROR(ABS(PazaryeriToplamSatis*AYARLAR!$B$29*(1-PazaryeriOrtKesintiOrani)),0)",
      sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "En Etkili Değişken", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 17, 2, "=IF(B14>=B15,\"Komisyon\",\"Satış Fiyatı\")", yazi=DIS_REF_YEŞIL, kalin=True)
    g = BarChart()
    g.type = "bar"
    g.title = "Tornado: Net Kâra Etki"
    g.add_data(Reference(ws, min_col=2, min_row=13, max_row=15), titles_from_data=True)
    g.set_categories(Reference(ws, min_col=1, min_row=14, max_row=15))
    g.height, g.width = 8, 13
    ws.add_chart(g, "E17")
    alt_bant(ws, 19, "Senaryo çarpanları ve tornado adımları AYARLAR'dan türetilir.")
    genislik(ws, {"A": 44, "B": 24, "C": 24, "D": 18})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", "0B1F3A", URUN_AD, son_kolon=40)
    ws.merge_cells("A1:L1")
    ws.merge_cells("A2:L2")
    h(ws, 1, 1, "PAZARYERİ NET KÂR VE EKSİK HAKEDİŞ — YÖNETİCİ ÖZETİ",
      kalin=True, boyut=16, yazi="FFFFFF", zemin=KOYU_LACIVERT, hiza="center")
    h(ws, 2, 1, "=_xlfn.CONCAT(\"Rapor Tarihi: \",TEXT(PazaryeriRaporTarihi,\"dd.mm.yyyy\"),\" | Hazırlayan: \",PazaryeriRaporHazirlayan,\" | Sürüm: \",PazaryeriDosyaSurumu)",
      yazi=GRİ, hiza="center", boyut=10)
    h(ws, 4, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 2, "=KARAR!B11", boyut=14, kalin=True, yazi=DIS_REF_YEŞIL)
    satirlar = [
        ("Toplam Satış (TL)", "=HESAP!B14", TL),
        ("Toplam Maliyet (TL)", "=HESAP!B19", TL),
        ("Toplam Net Kâr (TL)", "=HESAP!B20", TL),
        ("Genel Kâr Marjı", "=HESAP!B21", YÜZDE),
        ("Toplam Beklenen Hakediş (TL)", "=HESAP!B23", TL),
        ("Toplam Hakediş Farkı (TL)", "=HESAP!B24", TL),
        ("Eksik Hakediş Kalem Sayısı", "=HESAP!B25", CATI),
        ("En İyi Pazaryeri", "=HESAP!B27", None),
        ("Kalite Skoru", "=KONTROLLER!B19", CATI),
    ]
    for i, (ad, form, sayi) in enumerate(satirlar, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Varsayımlar ve Önemli Notlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    notlar = [
        "Beklenen hakediş, satış bedelinden komisyon, kargo/operasyon ve reklam kesintileri düşüldükten sonra pazaryerinin ödemesi gereken tutardır.",
        "Eksik hakediş tespiti, gerçek hakediş (pazaryerinden gelen ödeme) ile beklenen tutarın farkının eşikleri aşmasıyla üretilir; gerçek hakediş girilmeyen kalemler BEKLİYOR işaretlenir.",
        "Net kâr; satıştan ürün maliyeti, komisyon, kargo, reklam, iade maliyeti ve ek giderler düşülerek hesaplanır.",
        "Senaryo ve tornado analizi satış fiyatı ve komisyon değişimlerinin kâra etkisini gösterir; çarpanlar AYARLAR'dadır.",
        "Bu dosya karar destek aracıdır; vergi ve muhasebe hesabı mali müşavir görüşü yerine geçmez.",
        "Demo veriler gerçek müşteri verisi değildir; kullanımdan önce temizlenmelidir.",
    ]
    for i, m in enumerate(notlar, 17):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=12)
    h(ws, 24, 1, "Önerilen Aksiyonlar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    for i in range(3):
        h(ws, 25 + i, 1, f"=PazaryeriAksiyon{i+1}_ACIKLAMA", yazi="333333", kaydir=True)
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
    sutunlar = ["Kalem Kodu", "Ürün Adı", "Pazaryeri", "Birim Satış Fiyatı (TL)", "Miktar (Adet)",
                "Birim Ürün Maliyeti (TL)", "Komisyon Oranı", "Kargo/Operasyon Oranı",
                "Reklam Oranı", "İade Oranı", "Birim Ek Gider (TL)",
                "Beklenen Hakediş (TL)", "NET KÂR (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Beklenen Hakediş (TL)": '=IF(tblOrnekKalem[[#This Row],[Miktar (Adet)]]="",0,tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]*(1-tblOrnekKalem[[#This Row],[Komisyon Oranı]]-tblOrnekKalem[[#This Row],[Kargo/Operasyon Oranı]]-tblOrnekKalem[[#This Row],[Reklam Oranı]]))',
        "NET KÂR (TL)": '=IF(tblOrnekKalem[[#This Row],[Miktar (Adet)]]="",0,tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]-(tblOrnekKalem[[#This Row],[Birim Ürün Maliyeti (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]+tblOrnekKalem[[#This Row],[Komisyon Oranı]]*tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]+tblOrnekKalem[[#This Row],[Kargo/Operasyon Oranı]]*tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]+tblOrnekKalem[[#This Row],[Reklam Oranı]]*tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]+tblOrnekKalem[[#This Row],[İade Oranı]]*tblOrnekKalem[[#This Row],[Birim Satış Fiyatı (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]*(1-tblOrnekKalem[[#This Row],[Komisyon Oranı]])+tblOrnekKalem[[#This Row],[Birim Ek Gider (TL)]]*tblOrnekKalem[[#This Row],[Miktar (Adet)]]))',
        "Açıklama": '=IF(tblOrnekKalem[[#This Row],[Ürün Adı]]="","",CONCATENATE("Örnek kalem: ",tblOrnekKalem[[#This Row],[Ürün Adı]]))',
    }
    tablo_ekle(ws, "tblOrnekKalem", "A5:N1004", sutunlar, formuller)
    ornek = [
        ("P-1001", "Kulaklık Bluetooth", "Trendyol", 899, 25, 420, 0.18, 0.04, 0.06, 0.05, 12),
        ("P-1002", "Powerbank 20.000mAh", "Hepsiburada", 549, 60, 260, 0.22, 0.05, 0.05, 0.04, 9),
        ("P-1003", "Akıllı Saat", "Trendyol", 1499, 18, 820, 0.20, 0.04, 0.08, 0.06, 18),
        ("P-1004", "Şarj Aleti 65W", "Amazon", 349, 120, 160, 0.24, 0.06, 0.07, 0.03, 6),
        ("P-1005", "Kablo USB-C", "N11", 129, 300, 55, 0.15, 0.08, 0.04, 0.06, 3),
        ("P-1006", "Mouse Kablosuz", "Amazon", 279, 80, 130, 0.25, 0.06, 0.09, 0.05, 5),
    ]
    giris_kolonlari = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
    for i, satir in enumerate(ornek, 6):
        for k, deger in enumerate(satir):
            h(ws, i, giris_kolonlari[k], deger)
        for kolon in (4, 6, 11, 12, 13):
            ws.cell(row=i, column=kolon).number_format = TL
        for kolon in (7, 8, 9, 10):
            ws.cell(row=i, column=kolon).number_format = YÜZDE
        ws.cell(row=i, column=5).number_format = CATI
    genislik(ws, {"A": 11, "B": 24, "C": 15, "D": 15, "E": 11, "F": 16, "G": 12, "H": 16,
                  "I": 11, "J": 10, "K": 14, "L": 18, "M": 16, "N": 24})
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
    h(ws, 6, 1, "Pazaryerleri", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, paz in enumerate(PAZARYERLERI, 7):
        h(ws, i, 1, paz, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Pazaryeri | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni pazaryeri için satır ekleyin")
    genislik(ws, {"A": 22})


def ayarlar(ws):
    sayfa_hazirla(ws, "AYARLAR", "696969", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Sistem Ayarları ve Oranlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Tüm sabitler bu sayfadadır. Eşikler, senaryo çarpanları ve oranları "
                "güncelleyin; değişen her değer tüm hesaplamalara anında yansır.", yazi=GRİ, kaydir=True)
    baslik_satiri(ws, 6, ["anahtar", "deger", "birim", "alt", "ust", "aciklama", "kaynak",
                          "yururluk_tarihi", "dogrulama_tarihi"])
    TARIH_D = date(2026, 8, 10)
    YURURLUK = date(2026, 8, 10)
    DOGRULAMA = date(2026, 8, 10)
    satirlar = [
        (7, "PazaryeriFirmaUnvan", "Örnek Firma A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "PazaryeriRaporTarihi", TARIH_D, "Tarih", None, None, "Raporun dayanak tarihi.", "Kullanıcı"),
        (9, "PazaryeriRaporHazirlayan", "Pazaryeri Operasyon Müdürü", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "PazaryeriDosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "PazaryeriEksikHakedisEsik", 0.01, "Oran", 0, 1, "Fark oranı eşiği; üzerindeki farklar eksik hakediş sayılır.", "Varsayım"),
        (12, "PazaryeriEksikHakedisTutarEsik", 50, "TL", 0, 1000000000, "Fark tutarının ihmal edilmeyeceği asgari TL eşiği.", "Varsayım"),
        (13, "PazaryeriNetKarMarjiHedef", 0.10, "Oran", 0, 1, "GÜÇLÜ kararı için genel marj hedefi.", "Varsayım"),
        (14, "PazaryeriNetKarMarjiKritik", 0.00, "Oran", -1, 1, "Marjın altına inmesi durumunda İNCELE kararı üretilir.", "Varsayım"),
        (15, "PazaryeriKargoVarsayilanOran", 0.05, "Oran", 0, 1, "Kargo/operasyon oranı varsayılanı (kalem girişinde öneri).", "Varsayım"),
        (16, "PazaryeriReklamVarsayilanOran", 0.05, "Oran", 0, 1, "Reklam oranı varsayılanı.", "Varsayım"),
        (17, "PazaryeriIadeVarsayilanOran", 0.05, "Oran", 0, 1, "İade oranı varsayılanı.", "Varsayım"),
        (18, "PazaryeriSenIyiSatis", 0.05, "Oran", -0.5, 0.5, "İyimser senaryo satış fiyatı çarpanı.", "Varsayım"),
        (19, "PazaryeriSenBazSatis", 0.00, "Oran", -0.5, 0.5, "Baz senaryo satış fiyatı çarpanı.", "Varsayım"),
        (20, "PazaryeriSenKotuSatis", -0.05, "Oran", -0.5, 0.5, "Kötümser senaryo satış fiyatı çarpanı.", "Varsayım"),
        (21, "PazaryeriSenKritikSatis", -0.10, "Oran", -0.5, 0.5, "Kritik senaryo satış fiyatı çarpanı.", "Varsayım"),
        (22, "PazaryeriTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (23, "PazaryeriKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (24, "PazaryeriKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (25, "PazaryeriToplamFarkEsik", 500, "TL", 0, 1000000000, "Toplam hakediş farkının kontrol edildiği eşik.", "Varsayım"),
        (26, "PazaryeriP90Oran", 0.90, "Oran", 0, 1, "Satış fiyatı dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (27, "PazaryeriTemelKolonSayisi", 7, "Adet", 1, 20, "Veri kalitesi skorunda temel kabul edilen zorunlu giriş kolonu sayısı.", "Varsayım"),
        (28, "PazaryeriTornadoKomisyonOran", 0.01, "Oran", 0, 0.1, "Komisyon oranı tornado adımı (+1 puan).", "Varsayım"),
        (29, "PazaryeriTornadoSatisOran", -0.05, "Oran", -0.5, 0, "Satış fiyatı tornado adımı (−%5).", "Varsayım"),
        (30, "PazaryeriYuksekToplamSatisEsik", 1000000, "TL", 0, 100000000000, "Yüksek toplam satış kontrol eşiği.", "Varsayım"),
        (31, "PazaryeriAksiyon1_ACIKLAMA",
         "Eksik hakediş tespit edildi; pazaryeri müşteri hizmetlerine fark dökümünü ileterek tutarın iadesini talep edin.", "Metin", None, None,
         "İNCELE kararında gösterilen ilk öneri.", "Varsayım"),
        (32, "PazaryeriAksiyon2_ACIKLAMA",
         "Genel marj hedefin üzerinde; satışı artırmak için kâr marjı yüksek kalemlerin reklam bütçesini öne çekin.", "Metin", None, None,
         "GÜÇLÜ kararında gösterilen ikinci öneri.", "Varsayım"),
        (33, "PazaryeriAksiyon3_ACIKLAMA",
         "Marj kritik altında; komisyon yüksek pazaryerlerinde fiyat/komisyon yapısını yeniden müzakere edin veya portföyü gözden geçirin.", "Metin", None, None,
         "İNCELE/NORMAL kararında gösterilen üçüncü öneri.", "Varsayım"),
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
        h(ws, satir, 1, f"PazaryeriPanelSatir_{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"HESAP/PANO'daki kalem bazlı grafiğin tablodaki veri satırı ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel satır numarası {i} | "
                   f"Neden önemli: Grafikleri besler | "
                   f"Doğru kullanım: Tablodaki satır sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 34, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 22, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "KALEM_LISTESI sayfasında sarı hücrelere satış kalemlerinizi girin. Satış fiyatı, "
         "miktar, ürün maliyeti, komisyon/kargo/reklam/iade oranları ve ek gider doldurulunca "
         "net kâr ve marj otomatik hesaplanır."),
        ("Eksik hakediş yakalama",
         "Her kalem için pazaryerinden gelen gerçek ödemeyi 'Gerçek Hakediş' sütununa girin. "
         "Beklenen hakedişten fark, eşiklerin üzerindeyse kalem 'EKSİK HAKEDİŞ' olarak işaretlenir."),
        ("Eksik hakediş mutabakatı",
         "Beklenen hakediş; satış tutarından komisyon, kargo/operasyon ve reklam tutarları "
         "düşülerek hesaplanır. İade maliyeti mutabakata doğrudan girmez; kâr ve marjda etkilidir. "
         "Gerçek hakedişi KDV hariç girin; fark, AYARLAR'daki eşikleri aşınca EKSİK HAKEDİŞ işaretlenir."),
        ("Veri kalite skoru",
         "KONTROLLER sayfası zorunlu giriş kolonlarının doluluk oranını ölçer ve 0-100 puan üretir. "
         "Düşük skorda PANO'daki tüm çıktılara DÜŞÜK GÜVEN uyarısı basılır; karar kalitesini artırmak "
         "için eksik alanları doldurun."),
        ("Karar ve aksiyon",
         "KARAR sayfası marj ve eksik hakediş durumuna göre GÜÇLÜ/NORMAL/İNCELE kararını "
         "gerekçesiyle üretir ve önerilen aksiyonları gösterir."),
        ("Oranlar ve senaryolar",
         "AYARLAR sayfasından marj hedefleri, eksik hakediş eşikleri, senaryo ve tornado "
         "çarpanlarını güncelleyin; PANO ve SENARYO_DUYARLILIK anında güncellenir."),
        ("Güven ve karar kalitesi",
         "Veri kalite skoru girişlerin tamlığını ölçer. Kalem girişi yokken karar VERİ YOK; "
         "negatif girişte DURDUR üretir. Verileriniz cihazınızdan çıkmaz; dosya tamamen çevrimdışıdır."),
        ("Uyumluluk",
         "Excel 2016-365 (Windows ve Mac), LibreOffice ve Google Sheets ile uyumludur. "
         "Makro içermez; formüller ayrıştırıcı ile denetlenmiştir."),
        ("Karar kuralları",
         "KARAR sayfasındaki sonuç şu kurallarla üretilir: (1) kalem girişi yoksa VERİ YOK; "
         "(2) negatif satış fiyatı veya maliyet varsa DURDUR; (3) genel kâr marjı hedefin "
         "üzerindeyse ve eksik hakediş yoksa GÜÇLÜ; (4) eksik hakediş tespit edilmişse veya "
         "marj kritik eşiğin altındaysa İNCELE; (5) diğer durumlarda NORMAL. Gerekçe, kararın "
         "altında makine tarafından üretilir ve hangi kuralın tetiklendiğini açıklar."),
    ]
    for i, (baslik, metin) in enumerate(bolumler, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        ws.row_dimensions[i].height = 56
    genislik(ws, {"A": 34, "B": 110})


def kosullu_bicimlendirme(wb):
    kirmizi = PatternFill("solid", start_color="FDE9E9", end_color="FDE9E9")
    yesil = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
    sari = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
    kirmizi_font = Font(name=FONT, color="B3261E", bold=True, size=10)
    yesil_font = Font(name=FONT, color="1F7A4D", bold=True, size=10)

    ws = wb["KARAR"]
    for deger, dolu, yazifont in (("VERİ YOK", yesil, yesil_font), ("DURDUR", kirmizi, kirmizi_font),
                                   ("İNCELE", kirmizi, kirmizi_font), ("GÜÇLÜ", yesil, yesil_font),
                                   ("NORMAL", sari, None)):
        ws.conditional_formatting.add("B11",
            CellIsRule(operator="equal", formula=[f'"{deger}"'], fill=dolu, font=yazifont))

    ws = wb["PANO"]
    for r in (4, 6):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"İNCELE"'], fill=kirmizi, font=kirmizi_font))
    for c in ("B", "C", "D", "E", "F", "G", "H", "I"):
        for r in range(11, 16):
            ws.conditional_formatting.add(f"{c}{r}",
                CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

    ws = wb["KALEM_LISTESI"]
    for c in ("D", "F", "R"):
        ws.conditional_formatting.add(f"{c}6:{c}1004",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("X6:X1004",
        CellIsRule(operator="equal", formula=['"EKSİK HAKEDİŞ"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("X6:X1004",
        CellIsRule(operator="equal", formula=['"BEKLİYOR"'], fill=sari))
    ws.conditional_formatting.add("V6:V1004",
        CellIsRule(operator="greaterThan", formula=["PazaryeriEksikHakedisTutarEsik"], fill=sari, font=kirmizi_font))

    ws = wb["HESAP"]
    for hucre in ("B20",):
        ws.conditional_formatting.add(hucre,
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B21",
        CellIsRule(operator="lessThan", formula=["PazaryeriNetKarMarjiKritik"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B24",
        CellIsRule(operator="greaterThan", formula=["PazaryeriToplamFarkEsik"], fill=kirmizi, font=kirmizi_font))
    for r in range(6, 11):
        ws.conditional_formatting.add(f"I{r}",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))

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
        CellIsRule(operator="equal", formula=['"İNCELE"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B4",
        CellIsRule(operator="notEqual", formula=['"İNCELE"'], fill=yesil, font=yesil_font))


def main(cikti_yolu=None):
    wb = Workbook()
    ilk = wb.active
    sayfalar = [
        (kapak, "KAPAK"), (hizli_baslangic, "HIZLI_BASLANGIC"),
        (kalem_listesi, "KALEM_LISTESI"), (hesap, "HESAP"),
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
    ad_ekle(wb, "PazaryeriToplamSatis", "HESAP!$B$14")
    ad_ekle(wb, "PazaryeriToplamKomisyon", "HESAP!$B$15")
    ad_ekle(wb, "PazaryeriToplamKargo", "HESAP!$B$16")
    ad_ekle(wb, "PazaryeriToplamReklam", "HESAP!$B$17")
    ad_ekle(wb, "PazaryeriToplamIade", "HESAP!$B$18")
    ad_ekle(wb, "PazaryeriToplamMaliyet", "HESAP!$B$19")
    ad_ekle(wb, "PazaryeriToplamNetKar", "HESAP!$B$20")
    ad_ekle(wb, "PazaryeriGenelMarj", "HESAP!$B$21")
    ad_ekle(wb, "PazaryeriOrtKesintiOrani", "HESAP!$B$22")
    ad_ekle(wb, "PazaryeriToplamBeklenenHakedis", "HESAP!$B$23")
    ad_ekle(wb, "PazaryeriToplamFark", "HESAP!$B$24")
    ad_ekle(wb, "PazaryeriEksikKalemSayisi", "HESAP!$B$25")
    ad_ekle(wb, "PazaryeriOrtMarj", "HESAP!$B$26")
    ad_ekle(wb, "PazaryeriEnIyiPazaryeri", "HESAP!$B$27")
    ad_ekle(wb, "PazaryeriEnIyiMarj", "HESAP!$B$28")
    ad_ekle(wb, "PazaryeriOrtalamaNetKar", "HESAP!$B$44")
    ad_ekle(wb, "PazaryeriNetKarTahmin", "HESAP!$B$45")
    ad_ekle(wb, "PazaryeriNetKarStd", "HESAP!$B$46")
    # Ad tanımları — AYARLAR
    ayar_satirlari = {
        "PazaryeriFirmaUnvan": 7, "PazaryeriRaporTarihi": 8,
        "PazaryeriRaporHazirlayan": 9, "PazaryeriDosyaSurumu": 10,
        "PazaryeriEksikHakedisEsik": 11, "PazaryeriEksikHakedisTutarEsik": 12,
        "PazaryeriNetKarMarjiHedef": 13, "PazaryeriNetKarMarjiKritik": 14,
        "PazaryeriKargoVarsayilanOran": 15, "PazaryeriReklamVarsayilanOran": 16,
        "PazaryeriIadeVarsayilanOran": 17,
        "PazaryeriSenIyiSatis": 18, "PazaryeriSenBazSatis": 19,
        "PazaryeriSenKotuSatis": 20, "PazaryeriSenKritikSatis": 21,
        "PazaryeriTahminCarpan": 22, "PazaryeriKaliteIyi": 23, "PazaryeriKaliteOrta": 24,
        "PazaryeriToplamFarkEsik": 25, "PazaryeriP90Oran": 26,
        "PazaryeriTemelKolonSayisi": 27, "PazaryeriTornadoKomisyonOran": 28,
        "PazaryeriTornadoSatisOran": 29, "PazaryeriYuksekToplamSatisEsik": 30,
    }
    for i in range(1, 13):
        ayar_satirlari[f"PazaryeriPanelSatir_{i}"] = 34 + i
    for i in range(3):
        ayar_satirlari[f"PazaryeriAksiyon{i+1}_ACIKLAMA"] = 31 + i
    for ad, satir in ayar_satirlari.items():
        ad_ekle(wb, ad, f"AYARLAR!$B${satir}")
    ad_ekle(wb, "PazaryeriToplamGiris", "KONTROLLER!$B$17")
    ad_ekle(wb, "PazaryeriDoluGiris", "KONTROLLER!$B$18")
    # Modül katmanı (D03/D04/D05)
    modul_satirlari = {
        "modulT1OrtMarj": 6, "modulT1OrtMarjYorum": 6,
        "modulT2OrtKomisyon": 7, "modulT2OrtKomisyonYorum": 7,
        "modulO1KarSapmasi": 8, "modulO1KarSapmasiYorum": 8,
        "modulO2EnYuksekKar": 9, "modulO2EnYuksekKarYorum": 9,
        "modulO3Senaryo": 10, "modulO3SenaryoYorum": 10,
        "modulO6Hhi": 11, "modulO6HhiYorum": 11,
        "modulO8KaliteSkor": 12, "modulO8KaliteSkorYorum": 12,
        "modulI1NetKarTahmin": 13, "modulI1NetKarTahminYorum": 13,
        "modulI2YuzdelikP90": 14, "modulI2YuzdelikP90Yorum": 14,
        "modulI8HakedisEtkisi": 15, "modulI8HakedisEtkisiYorum": 15,
    }
    for ad, satir in modul_satirlari.items():
        kolon = "D" if ad.endswith("Yorum") else "C"
        ad_ekle(wb, ad, f"ANALITIK_MOTOR!${kolon}${satir}")
    ad_ekle(wb, "ListePazaryerleri", "LISTELER!$A$7:$A$11")

    kosullu_bicimlendirme(wb)

    for wsx in wb.worksheets:
        sayfa_koru(wsx)

    tablo_formullerini_hucrelere_yaz(wb, satir_basi=6, satir_sonu=1004)

    wb.calculation.fullCalcOnLoad = True
    dosya = "PazaryeriNetKarVeEksikHakedisYakalayici.xlsx"
    wb.save(cikti_yolu or dosya)
    if not cikti_yolu:
        print(f"Dosya oluşturuldu: {dosya}")
    return cikti_yolu or dosya


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
