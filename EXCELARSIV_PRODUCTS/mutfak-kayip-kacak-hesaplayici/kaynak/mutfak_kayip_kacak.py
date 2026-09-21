"""
Mutfak Kayıp/Kaçak Hesaplayıcı — üretim betiği.
Manda v4 uyumlu: sayfa sırası, sarı giriş, 1234 koruma, ipucu/not, tablo mimarisi.
Teorik tüketim (satış adedi × reçete) ile fiili tüketimi (stok hareketi) karşılaştırır.
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

URUN_AD = "Mutfak Kayıp/Kaçak Hesaplayıcı"
SURUM = "1.0.0"
RENK = "B3261E"

BIRIMLER = ["kg", "gr", "lt", "adet", "paket", "m"]
HAMMADDELER = ["Un", "Yumurta", "Tereyağı", "Peynir", "Zeytin", "Salça", "Et", "Tavuk", "Patates", "Domates", "Diğer"]
URUN_KODLARI = ["UR-01", "UR-02", "UR-03", "UR-04", "UR-05", "UR-06"]


def kapak(ws):
    sayfa_hazirla(ws, "KAPAK", RENK, URUN_AD, son_kolon=40)
    bant(ws, 3, URUN_AD, boyut=18, son_kolon=40)
    h(ws, 4, 1, "Satılan ürünlerin reçetelerinden üretilen teorik hammadde tüketimini stok "
                "hareketiyle karşılaştırarak mutfaktaki kayıp ve kaçağı tutar ve oran olarak "
                "tespit edin; kritik kalemleri gerekçesiyle raporlayın.", kaydir=True, yazi=GRİ)
    ws.merge_cells("A4:P4")
    h(ws, 6, 1, "Bu ürün ne sağlar?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    faydalar = [
        "Ürün reçetesi ve satış adedi üzerinden teorik hammadde tüketiminin otomatik hesaplanması",
        "Dönem başı stok + alım − dönem sonu stok ile fiili tüketimin ve kayıp miktarının bulunması",
        "Hammadde bazında kayıp tutarı, kayıp oranı ve toplam kayıp/kaçak maliyetinin hesaplanması",
        "Kayıp yoğunlaşması, anomali, senaryo ve duyarlılık analizi ile karar kapısı ve aksiyon önerisi",
        "Kritik kayıp kalemlerinin KONTROL ALTINDA / İNCELE / KRİTİK KAYIP kararıyla önceliklendirilmesi",
    ]
    for i, m in enumerate(faydalar, 7):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 14, 1, "Kimler için?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    kimler = [
        "Kayıp ve kaçağı ölçmek isteyen restoran, kafe ve toplu yemek işletmeleri",
        "Stok ile satışı mutabakat etmek isteyen aşçıbaşı, mutfak şefi ve işletme sahipleri",
        "Kâr marjını düşüren kayıp kalemlerini kanıtla raporlamak isteyen finans yöneticileri",
        "Müşterisine stok ve kayıp kontrolü analizi öneren mali müşavirler",
    ]
    for i, m in enumerate(kimler, 15):
        h(ws, i, 1, "• " + m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
    h(ws, 20, 1, "Nasıl başlanır?", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 21, 1, "1. HIZLI_BASLANGIC sayfasını okuyun. "
                 "2. URUN_LISTESI, RECETE ve STOK sayfalarına verilerinizi girin. "
                 "3. PANO ve KARAR sayfalarından kayıp/kaçak kararını izleyin.", yazi="333333", kaydir=True)
    ws.merge_cells("A21:P21")
    h(ws, 23, 1, "Sürüm " + SURUM + " | 2026 | ExcelArşiv | Lisans: Tek kullanıcı",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A23:P23")
    genislik(ws, {"A": 60})


def hizli_baslangic(ws):
    sayfa_hazirla(ws, "HIZLI_BASLANGIC", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Nasıl kullanılır?", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    adimlar = [
        ("Adım 1 — Ürün listesini girin",
         "URUN_LISTESI sayfasındaki sarı hücrelere ürün kodunu/adını, satış adedini ve birim "
         "satış fiyatını girin. Reçete maliyeti ve hasılat otomatik hesaplanır."),
        ("Adım 2 — Reçeteleri tanımlayın",
         "RECETE sayfasındaki sarı hücrelere her ürün için hammadde, birim, miktar ve birim "
         "maliyeti girin; teorik tüketim otomatik toplanır."),
        ("Adım 3 — Stok hareketini girin",
         "STOK sayfasında dönem başı, alım ve dönem sonu stokları girin; fiili tüketim ve "
         "kayıp/kaçak otomatik hesaplanır."),
        ("Adım 4 — Kararı değerlendirin",
         "KARAR sayfası kayıp oranını eşiklerle karşılaştırarak KONTROL ALTINDA / İNCELE / "
         "KRİTİK KAYIP kararını gerekçesiyle üretir."),
    ]
    for i, (baslik, metin) in enumerate(adimlar, 5):
        h(ws, i, 1, baslik, kalin=True, yazi=KOYU_LACIVERT, kaydir=True)
        h(ws, i, 2, metin, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=20)
        yorum_ekle(ws, f"A{i}", f"Tanım: {baslik} | Neden önemli: Ürünü doğru kullanmak için | "
                                f"Doğru kullanım: Sırayla izleyin")
    h(ws, 11, 1, "Sık yapılan hatalar", kalin=True, boyut=12, yazi="B3261E")
    hatalar = [
        "Kayıp oranını yüzde olarak 5 değil 0,05 girmek",
        "Dönem sonu stoku boş bırakmak (fiili tüketim hatalı çıkar)",
        "Stok birimleri ile reçete birimlerini karıştırmak (kg/gr uyumsuzluğu)",
    ]
    for i, m in enumerate(hatalar, 12):
        h(ws, i, 1, "• " + m, yazi="B3261E")
    h(ws, 16, 1, "Sık sorulan sorular", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    sss = [
        "Soru: Kayıp oranını nasıl girerim? Cevap: Oran girilmez; satış adedi, reçete ve stok verisinden otomatik hesaplanır.",
        "Soru: Fiili tüketim nasıl bulunur? Cevap: Dönem başı stok + alım − dönem sonu stok formülüyle hesaplanır.",
        "Soru: Teorik tüketim nasıl bulunur? Cevap: Her ürünün satış adedi, reçetedeki miktarla çarpılıp hammadde bazında toplanır.",
        "Soru: Hammadde listesinde aradığım yoksa? Cevap: LISTELER sayfasındaki sarı hücreye yeni hammadde adını ekleyin.",
    ]
    for i, m in enumerate(sss, 17):
        h(ws, i, 1, m, yazi="333333", kaydir=True)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=30)
        ws.row_dimensions[i].height = 32
    h(ws, 22, 1, "Bu dosya karar destek aracıdır; kayıp tespiti sayım ve muhasebe kayıtlarıyla "
                 "doğrulanmalıdır.", yazi=GRİ, boyut=9, kaydir=True)
    genislik(ws, {"A": 70, "B": 60})


def urun_listesi(ws):
    sayfa_hazirla(ws, "URUN_LISTESI", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Ürün Listesi ve Satışlar", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Sarı hücrelere manuel veri girin. Reçete maliyeti, hasılat ve kâr katkısı "
                "otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ürün Kodu", "Ürün Adı", "Satış Adedi (Adet)", "Birim Satış Fiyatı (TL)",
                "Reçete Maliyeti (TL)", "Satış Hasılatı (TL)", "Kâr Katkısı (TL)",
                "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Reçete Maliyeti (TL)": '=IF(tblUrun[[#This Row],[Ürün Kodu]]="","",SUMIFS(tblRecete[Tutar (TL)],tblRecete[Ürün Kodu],tblUrun[[#This Row],[Ürün Kodu]]))',
        "Satış Hasılatı (TL)": '=IF(tblUrun[[#This Row],[Ürün Kodu]]="","",tblUrun[[#This Row],[Satış Adedi (Adet)]]*tblUrun[[#This Row],[Birim Satış Fiyatı (TL)]])',
        "Kâr Katkısı (TL)": '=IF(tblUrun[[#This Row],[Ürün Kodu]]="","",tblUrun[[#This Row],[Satış Hasılatı (TL)]]-tblUrun[[#This Row],[Reçete Maliyeti (TL)]])',
        "Açıklama": '=IF(tblUrun[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblUrun[[#This Row],[Ürün Adı]]," | Hasılat: ",TEXT(tblUrun[[#This Row],[Satış Hasılatı (TL)]],"₺ #,##0")," | Kâr Katkısı: ",TEXT(tblUrun[[#This Row],[Kâr Katkısı (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblUrun", "A5:H1004", sutunlar, formuller)
    dogrulama(ws, "decimal", "0", "C6:C1004",
              baslik="Satış Adedi", mesaj="Dönemde satılan adedi girin.",
              hata_baslik="Geçersiz adet", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "decimal", "0", "D6:D1004",
              baslik="Birim Satış Fiyatı", mesaj="Birim satış fiyatını TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 10.000.000 arasında olmalı.",
              isaret="between", f2="10000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = CATI
        for kolon in (4, 5, 6, 7):
            ws.cell(row=satir, column=kolon).number_format = TL
    ornek = [
        ("UR-01", "Klasik Pide", 320, 120),
        ("UR-02", "Lahmacun", 410, 95),
        ("UR-03", "Kumpir", 280, 140),
        ("UR-04", "Tost Sandviç", 520, 75),
        ("UR-05", "Çiğ Köfte Menü", 190, 150),
        ("UR-06", "İskender", 230, 180),
    ]
    for i, (kod, urun, adet, fiyat) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = kod
        ws.cell(row=i, column=2).value = urun
        ws.cell(row=i, column=3).value = adet
        ws.cell(row=i, column=4).value = fiyat
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4])
    sabitle(ws, "A6")
    genislik(ws, {"A": 11, "B": 20, "C": 16, "D": 18, "E": 18, "F": 18, "G": 16, "H": 52})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def recete(ws):
    sayfa_hazirla(ws, "RECETE", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Ürün Reçeteleri", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Her ürünün hammadde ihtiyacını girin; teorik tüketim stok sayfasına "
                "aktarılır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Ürün Kodu", "Ürün Adı", "Hammadde", "Birim", "Miktar",
                "Birim Maliyet (TL)", "Tutar (TL)", "Satış Adedi (Adet)",
                "Teorik Tüketim", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Tutar (TL)": '=IF(tblRecete[[#This Row],[Miktar]]="","",tblRecete[[#This Row],[Miktar]]*tblRecete[[#This Row],[Birim Maliyet (TL)]])',
        "Satış Adedi (Adet)": '=IF(tblRecete[[#This Row],[Ürün Kodu]]="","",IFERROR(INDEX(tblUrun[Satış Adedi (Adet)],MATCH(tblRecete[[#This Row],[Ürün Kodu]],tblUrun[Ürün Kodu],0)),0))',
        "Teorik Tüketim": '=IF(tblRecete[[#This Row],[Miktar]]="","",tblRecete[[#This Row],[Miktar]]*tblRecete[[#This Row],[Satış Adedi (Adet)]])',
        "Açıklama": '=IF(tblRecete[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblRecete[[#This Row],[Ürün Adı]]," | Hammadde: ",tblRecete[[#This Row],[Hammadde]]," | Teorik: ",TEXT(tblRecete[[#This Row],[Teorik Tüketim]],"#,##0.000")))',
    }
    tablo_ekle(ws, "tblRecete", "A5:J1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeHammaddeler", "C6:C1004",
              baslik="Hammadde", mesaj="Listeden hammaddeyi seçin.",
              hata_baslik="Geçersiz hammadde", hata_mesaj="Listede olmayan hammadde giremezsiniz.")
    dogrulama(ws, "list", "ListeBirimler", "D6:D1004",
              baslik="Birim", mesaj="Listeden birimi seçin.",
              hata_baslik="Geçersiz birim", hata_mesaj="Listede olmayan birim giremezsiniz.")
    dogrulama(ws, "decimal", "0", "E6:E1004",
              baslik="Miktar", mesaj="Reçetedeki hammadde miktarını girin.",
              hata_baslik="Geçersiz miktar", hata_mesaj="0 veya daha büyük olmalı.")
    dogrulama(ws, "decimal", "0", "F6:F1004",
              baslik="Birim Maliyet", mesaj="Hammadde birim maliyetini TL olarak girin.",
              hata_baslik="Geçersiz değer", hata_mesaj="0 ile 1.000.000 arasında olmalı.",
              isaret="between", f2="1000000")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=5).number_format = CATI
        for kolon in (6, 7):
            ws.cell(row=satir, column=kolon).number_format = TL
        ws.cell(row=satir, column=8).number_format = CATI
        ws.cell(row=satir, column=9).number_format = CATI
    ornek = [
        ("UR-01", "Klasik Pide", "Un", "kg", 0.250, 18),
        ("UR-01", "Klasik Pide", "Tereyağı", "kg", 0.030, 120),
        ("UR-01", "Klasik Pide", "Peynir", "kg", 0.080, 160),
        ("UR-01", "Klasik Pide", "Yumurta", "adet", 1, 6),
        ("UR-02", "Lahmacun", "Un", "kg", 0.080, 18),
        ("UR-02", "Lahmacun", "Et", "kg", 0.040, 320),
        ("UR-02", "Lahmacun", "Domates", "kg", 0.050, 35),
        ("UR-02", "Lahmacun", "Salça", "kg", 0.010, 60),
        ("UR-03", "Kumpir", "Patates", "kg", 0.350, 22),
        ("UR-03", "Kumpir", "Tereyağı", "kg", 0.040, 120),
        ("UR-03", "Kumpir", "Peynir", "kg", 0.050, 160),
        ("UR-04", "Tost Sandviç", "Un", "kg", 0.120, 18),
        ("UR-04", "Tost Sandviç", "Peynir", "kg", 0.060, 160),
        ("UR-04", "Tost Sandviç", "Tereyağı", "kg", 0.020, 120),
        ("UR-05", "Çiğ Köfte Menü", "Un", "kg", 0.100, 18),
        ("UR-05", "Çiğ Köfte Menü", "Domates", "kg", 0.040, 35),
        ("UR-05", "Çiğ Köfte Menü", "Salça", "kg", 0.020, 60),
        ("UR-06", "İskender", "Et", "kg", 0.180, 320),
        ("UR-06", "İskender", "Tereyağı", "kg", 0.030, 120),
        ("UR-06", "İskender", "Patates", "kg", 0.150, 22),
    ]
    for i, (kod, urun, ham, birim, miktar, maliyet) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = kod
        ws.cell(row=i, column=2).value = urun
        ws.cell(row=i, column=3).value = ham
        ws.cell(row=i, column=4).value = birim
        ws.cell(row=i, column=5).value = miktar
        ws.cell(row=i, column=6).value = maliyet
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 6])
    sabitle(ws, "A6")
    genislik(ws, {"A": 11, "B": 20, "C": 16, "D": 8, "E": 9, "F": 15, "G": 13, "H": 15,
                  "I": 13, "J": 52})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def stok(ws):
    sayfa_hazirla(ws, "STOK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Hammadde Stok Hareketi", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Dönem başı stok, alım ve dönem sonu stok girin; fiili tüketim ve kayıp/kaçak "
                "otomatik hesaplanır.", yazi=GRİ, kaydir=True)
    sutunlar = ["Hammadde", "Birim", "Dönem Başı Stok", "Alım", "Dönem Sonu Stok",
                "Fiili Tüketim", "Birim Maliyet (TL)", "Teorik Tüketim",
                "Kayıp/Kaçak (Miktar)", "Kayıp Oranı", "Kayıp Tutarı (TL)", "Açıklama"]
    baslik_satiri(ws, 5, sutunlar)
    formuller = {
        "Fiili Tüketim": '=IF(tblStok[[#This Row],[Hammadde]]="","",MAX(0,tblStok[[#This Row],[Dönem Başı Stok]]+tblStok[[#This Row],[Alım]]-tblStok[[#This Row],[Dönem Sonu Stok]]))',
        "Teorik Tüketim": '=IF(tblStok[[#This Row],[Hammadde]]="","",SUMIFS(tblRecete[Teorik Tüketim],tblRecete[Hammadde],tblStok[[#This Row],[Hammadde]]))',
        "Kayıp/Kaçak (Miktar)": '=IF(tblStok[[#This Row],[Hammadde]]="","",MAX(0,tblStok[[#This Row],[Fiili Tüketim]]-tblStok[[#This Row],[Teorik Tüketim]]))',
        "Kayıp Oranı": '=IF(tblStok[[#This Row],[Hammadde]]="","",IF(tblStok[[#This Row],[Teorik Tüketim]]=0,0,tblStok[[#This Row],[Kayıp/Kaçak (Miktar)]]/tblStok[[#This Row],[Teorik Tüketim]]))',
        "Kayıp Tutarı (TL)": '=IF(tblStok[[#This Row],[Hammadde]]="","",tblStok[[#This Row],[Kayıp/Kaçak (Miktar)]]*tblStok[[#This Row],[Birim Maliyet (TL)]])',
        "Açıklama": '=IF(tblStok[[#This Row],[Hammadde]]="","",CONCATENATE("Hammadde: ",tblStok[[#This Row],[Hammadde]]," | Kayıp: ",TEXT(tblStok[[#This Row],[Kayıp Tutarı (TL)]],"₺ #,##0")," | Oran: ",TEXT(tblStok[[#This Row],[Kayıp Oranı]],"0,0%")))',
    }
    tablo_ekle(ws, "tblStok", "A5:L1004", sutunlar, formuller)
    dogrulama(ws, "list", "ListeHammaddeler", "A6:A1004",
              baslik="Hammadde", mesaj="Listeden hammaddeyi seçin.",
              hata_baslik="Geçersiz hammadde", hata_mesaj="Listede olmayan hammadde giremezsiniz.")
    dogrulama(ws, "list", "ListeBirimler", "B6:B1004",
              baslik="Birim", mesaj="Listeden birimi seçin.",
              hata_baslik="Geçersiz birim", hata_mesaj="Listede olmayan birim giremezsiniz.")
    for kolon in ("C", "D", "E", "G"):
        dogrulama(ws, "decimal", "0", f"{kolon}6:{kolon}1004",
                  baslik="Stok değeri", mesaj="0 veya daha büyük değer girin.",
                  hata_baslik="Geçersiz değer", hata_mesaj="0 veya daha büyük olmalı.")
    for satir in range(6, 1004):
        ws.cell(row=satir, column=3).number_format = CATI
        for kolon in (4, 5, 6, 8, 9):
            ws.cell(row=satir, column=kolon).number_format = CATI
        ws.cell(row=satir, column=7).number_format = TL
        ws.cell(row=satir, column=10).number_format = YÜZDE
        ws.cell(row=satir, column=11).number_format = TL
    ornek = [
        ("Un", "kg", 120, 90, 135),
        ("Tereyağı", "kg", 18, 14, 12),
        ("Peynir", "kg", 30, 25, 20),
        ("Yumurta", "adet", 400, 300, 350),
        ("Et", "kg", 60, 55, 40),
        ("Domates", "kg", 45, 40, 30),
        ("Salça", "kg", 12, 10, 8),
        ("Patates", "kg", 100, 80, 70),
    ]
    birim_maliyet = {"Un": 18, "Tereyağı": 120, "Peynir": 160, "Yumurta": 6,
                     "Et": 320, "Domates": 35, "Salça": 60, "Patates": 22}
    for i, (ham, birim, bas, alim, son) in enumerate(ornek, 6):
        ws.cell(row=i, column=1).value = ham
        ws.cell(row=i, column=2).value = birim
        ws.cell(row=i, column=3).value = bas
        ws.cell(row=i, column=4).value = alim
        ws.cell(row=i, column=5).value = son
        ws.cell(row=i, column=7).value = birim_maliyet[ham]
    giris_hucreleri(ws, 6, 1004, [1, 2, 3, 4, 5, 7])
    sabitle(ws, "A6")
    genislik(ws, {"A": 16, "B": 8, "C": 15, "D": 10, "E": 15, "F": 13, "G": 15, "H": 14,
                  "I": 17, "J": 11, "K": 16, "L": 52})
    alt_bant(ws, 1006, "Sarı hücreler manuel giriştir; formül hücreleri kilitli ve korumalıdır.")


def hesap(ws):
    sayfa_hazirla(ws, "HESAP", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kayıp/Kaçak Hesaplama Motoru", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Stok ve reçete verilerinden canlı üretilir; toplam kayıp, kayıp oranı ve "
                "kritik kalemler hesaplanır.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "TOPLAM HASILAT (TL)", yazi="333333", kalin=True)
    h(ws, 6, 2, "=SUM(tblUrun[Satış Hasılatı (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 7, 1, "TOPLAM REÇETE MALİYETİ (TL)", yazi="333333", kalin=True)
    h(ws, 7, 2, "=SUM(tblUrun[Reçete Maliyeti (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 8, 1, "Toplam Kâr Katkısı (TL)", yazi="333333")
    h(ws, 8, 2, "=SUM(tblUrun[Kâr Katkısı (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Toplam Fiili Tüketim (TL)", yazi="333333")
    h(ws, 9, 2, "=SUMPRODUCT(tblStok[Fiili Tüketim],tblStok[Birim Maliyet (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "Toplam Teorik Tüketim (TL)", yazi="333333")
    h(ws, 10, 2, "=SUMPRODUCT(tblStok[Teorik Tüketim],tblStok[Birim Maliyet (TL)])", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "TOPLAM KAYIP/KAÇAK TUTARI (TL)", yazi="333333", kalin=True)
    h(ws, 11, 2, "=IFERROR(B9-B10,0)", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 12, 1, "Toplam Kayıp Oranı", yazi="333333")
    h(ws, 12, 2, "=IFERROR(B11/B10,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 13, 1, "Kayıp Oranı Eşiği (İncele)", yazi="333333")
    h(ws, 13, 2, "=IF(COUNTA(tblStok[Hammadde])>0,MutfakInceleEsik,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 14, 1, "Kayıp Oranı Eşiği (Kritik)", yazi="333333")
    h(ws, 14, 2, "=IF(COUNTA(tblStok[Hammadde])>0,MutfakKritikEsik,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 15, 1, "Kayıp Hammadde Sayısı", yazi="333333")
    h(ws, 15, 2, '=COUNTIF(tblStok[Kayıp/Kaçak (Miktar)],">0")', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "En Kritik Hammadde", yazi="333333")
    h(ws, 16, 2, '=IFERROR(INDEX(tblStok[Hammadde],MATCH(MAX(tblStok[Kayıp Tutarı (TL)]),tblStok[Kayıp Tutarı (TL)],0)),"")',
      yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "En Yüksek Kayıp Tutarı (TL)", yazi="333333")
    h(ws, 17, 2, '=IFERROR(MAX(tblStok[Kayıp Tutarı (TL)]),0)', sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Kayıp Stok Değerine Oranı", yazi="333333")
    h(ws, 18, 2, "=IFERROR(SUMPRODUCT(tblStok[Dönem Başı Stok],tblStok[Birim Maliyet (TL)])+SUMPRODUCT(tblStok[Alım],tblStok[Birim Maliyet (TL)]),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 19, 1, "Kâr Katkısı Kayıp Sonrası (TL)", yazi="333333", kalin=True)
    h(ws, 19, 2, "=B8-B11", sayi=TL, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 20, 1, "Kâr Katkısı Kayıp Oranı", yazi="333333")
    h(ws, 20, 2, "=IFERROR(B11/B8,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    # Panel verisi (grafikler için) — ürün sırasına göre
    h(ws, 23, 1, "Ürün Sırası", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 23, 2, "Hasılat (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 23, 4, "Reçete Maliyeti (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 23, 6, "Kâr Katkısı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(6):
        h(ws, 24 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 24 + i, 2, f"=IFERROR(INDEX(tblUrun[Satış Hasılatı (TL)],MutfakPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 24 + i, 4, f"=IFERROR(INDEX(tblUrun[Reçete Maliyeti (TL)],MutfakPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, 24 + i, 6, f"=IFERROR(INDEX(tblUrun[Kâr Katkısı (TL)],MutfakPanelSatir_{i+1}),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    # Projeksiyon yardımcı serisi (kayıp tutarı trendi)
    h(ws, 32, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 32, 2, "Kayıp Tutarı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    kayip_tutarlari = [("1", "=IFERROR(INDEX(tblStok[Kayıp Tutarı (TL)],1),0)"),
                       ("2", "=IFERROR(INDEX(tblStok[Kayıp Tutarı (TL)],2),0)"),
                       ("3", "=IFERROR(INDEX(tblStok[Kayıp Tutarı (TL)],3),0)"),
                       ("4", "=IFERROR(INDEX(tblStok[Kayıp Tutarı (TL)],4),0)")]
    for i, (donem, form) in enumerate(kayip_tutarlari, 33):
        h(ws, i, 1, donem, sayi=CATI, yazi=GRİ)
        h(ws, i, 2, form, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 38, 1, "Ortalama Kayıp Tutarı (TL)", yazi="333333")
    h(ws, 38, 2, "=IFERROR(AVERAGE(B33:B36),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 39, 1, "Gelecek Dönem Kayıp Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 39, 2, "=IFERROR(FORECAST.LINEAR(MutfakTahminDonem,B33:B36,MutfakProjX)-B36,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 40, 1, "Kayıp Std Sapması (TL)", yazi="333333")
    h(ws, 40, 2, "=IFERROR(STDEV.P(B33:B36),0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 41, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 41, 2, "=IFERROR(B38-B40*MutfakTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 42, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 42, 2, "=IFERROR(B38+B40*MutfakTahminCarpan,0)", sayi=TL, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 44, "Motor, stok ve reçete verilerinden canlı beslenir; tüm eşikler AYARLAR'dan okunur.")
    genislik(ws, {"A": 34, "B": 40, "C": 24, "D": 24, "E": 24, "F": 22})


def kontroller(ws):
    sayfa_hazirla(ws, "KONTROLLER", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Veri Kalite ve Kontrol Merkezi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Girdilerin kalitesini ölçer; eksik, negatif ve uç değerleri raporlar.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    baslik_satiri(ws, 5, ["Kontrol", "Sonuç", "Durum"])
    kontroller = [
        ("Ürün girildi mi?", '=IF(COUNTA(tblUrun[Ürün Adı])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Reçete kalemi girildi mi?", '=IF(COUNTA(tblRecete[Ürün Adı])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Stok satırı girildi mi?", '=IF(COUNTA(tblStok[Hammadde])=0,"BOŞ","DOLU")', NORMAL, "B3261E"),
        ("Negatif stok girişi var mı?", '=IF(COUNTIF(tblStok[Dönem Başı Stok],"<0")+COUNTIF(tblStok[Alım],"<0")+COUNTIF(tblStok[Dönem Sonu Stok],"<0")=0,"TEMİZ","NEGATİF VAR")', NORMAL, "B3261E"),
        ("Kayıp oranı 0,50 üzerinde mi?", '=IF(COUNTIF(tblStok[Kayıp Oranı],">0.5")=0,"TEMİZ","ORAN >0,5 VAR")', NORMAL, "B3261E"),
        ("Ürün satış adedi girildi mi?", '=IF(COUNTIF(tblUrun[Satış Adedi (Adet)],">0")>=3,"YETERLİ","AZ")', NORMAL, "B3261E"),
    ]
    for i, (ad, form, durum1, durum2) in enumerate(kontroller, 6):
        h(ws, i, 1, ad, yazi="333333")
        h(ws, i, 2, form, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, "AUTOMATİK", yazi=GRİ)
    h(ws, 14, 1, "Veri Kalite Skoru", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 15, 1, "Doluluk Oranı (%)", yazi="333333")
    h(ws, 15, 2, '=IFERROR((COUNTA(tblUrun[Ürün Adı])+COUNTA(tblRecete[Ürün Adı])+COUNTA(tblStok[Hammadde]))/(COUNTA(tblUrun[Ürün Adı])+COUNTA(tblRecete[Ürün Adı])+COUNTA(tblStok[Hammadde])+COUNTBLANK(tblUrun[Ürün Adı])+COUNTBLANK(tblRecete[Ürün Adı])+COUNTBLANK(tblStok[Hammadde]))*100,0)',
      sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 16, 1, "Veri Kalite Skoru", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 16, 2, '=MIN(100,MAX(0,B15))', sayi=CATI, yazi=DIS_REF_YEŞIL)
    h(ws, 17, 1, "Kalite Durumu", yazi="333333")
    h(ws, 17, 2, '=IF(B16>=MutfakKaliteIyi,"YÜKSEK",IF(B16>=MutfakKaliteOrta,"ORTA","DÜŞÜK"))', yazi=DIS_REF_YEŞIL)
    h(ws, 18, 1, "Canlı Formül Sayacı", yazi="333333")
    h(ws, 18, 2, '=IF(COUNTIF(tblStok[Açıklama],"")=0,COUNTA(tblStok[Açıklama]),0)', sayi=CATI, yazi=DIS_REF_YEŞIL)
    alt_bant(ws, 20, "Tüm kontroller canlı formüllerle üretilir; sahte geçiş yoktur.")
    genislik(ws, {"A": 38, "B": 60, "C": 14})


def karar(ws):
    sayfa_hazirla(ws, "KARAR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kayıp/Kaçak Karar Kapısı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Kayıp oranı, kayıp tutarı ve stok değerine göre karar üretir.",
      yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Toplam Kayıp Oranı", yazi="333333")
    h(ws, 6, 2, "=HESAP!B12", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Toplam Kayıp Tutarı (TL)", yazi="333333")
    h(ws, 7, 2, "=HESAP!B11", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 8, 1, "İncele Eşiği", yazi="333333")
    h(ws, 8, 2, "=IF(COUNTA(tblStok[Hammadde])>0,MutfakInceleEsik,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 9, 1, "Kritik Eşiği", yazi="333333")
    h(ws, 9, 2, "=IF(COUNTA(tblStok[Hammadde])>0,MutfakKritikEsik,0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    h(ws, 10, 1, "En Kritik Hammadde", yazi="333333")
    h(ws, 10, 2, "=HESAP!B16", yazi=DIS_REF_YEŞIL)
    h(ws, 11, 1, "KARAR", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 11, 2, '=IF(COUNTA(tblStok[Hammadde])=0,"VERİ YOK",'
                 'IF(OR(COUNTIF(tblStok[Dönem Başı Stok],"<0")>0,COUNTIF(tblStok[Alım],"<0")>0,COUNTIF(tblStok[Dönem Sonu Stok],"<0")>0),"DURDUR",'
                 'IF(HESAP!B12>=MutfakKritikEsik,"KRİTİK KAYIP",'
                 'IF(AND(HESAP!B12>0,HESAP!B12<MutfakKritikEsik,HESAP!B12>=MutfakInceleEsik),"İNCELE",'
                 'IF(AND(HESAP!B12>0,HESAP!B12<MutfakInceleEsik),"DÜŞÜK KAYIP","KONTROL ALTINDA")))))',
      boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 12, 2, '=IF(B11="VERİ YOK","Stok veya ürün verisi girilmedi; karar üretilemiyor.",'
                 'IF(B11="DURDUR","Negatif stok girişi tespit edildi; veriler düzeltilmeli.",'
                 'IF(B11="KRİTİK KAYIP","Toplam kayıp oranı kritik eşiğin üzerinde; acil müdahale gereklidir.",'
                 'IF(B11="İNCELE","Kayıp oranı inceleme aralığında; neden analizi yapılmalıdır.",'
                 'IF(B11="DÜŞÜK KAYIP","Kayıp oranı düşük; süreç kontrolüne devam edilebilir.",'
                 '"Kayıp tespit edilmedi; stok ve satış mutabakatı tutarlıdır.")))))',
      yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[12].height = 44
    h(ws, 14, 1, "Önerilen Aksiyonlar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    for i, ad in enumerate(["Aksiyon 1", "Aksiyon 2", "Aksiyon 3"], 15):
        h(ws, i, 1, ad, yazi="333333", kalin=True)
        h(ws, i, 2, f'=IF(B11="VERİ YOK","Ürün, reçete ve stok verilerini girin; karar otomatik üretilir.",'
                    f'IF(B11="DURDUR","Negatif stok değerlerini düzeltin.",MutfakAksiyon{i-14}_ACIKLAMA))',
          yazi=DIS_REF_YEŞIL, kaydir=True)
        ws.row_dimensions[i].height = 40
    h(ws, 19, 1, "Karar Notu", yazi="333333")
    h(ws, 19, 2, "Bu karar destek aracıdır; kayıp tespiti fiili sayımla doğrulanmalıdır.", yazi=GRİ)
    baski_hazirla(ws, "A1:B19", URUN_AD)
    alt_bant(ws, 21, "Karar kuralları KILAVUZ sayfasında gösterilmiştir.")
    genislik(ws, {"A": 32, "B": 80})


def pano(ws):
    sayfa_hazirla(ws, "PANO", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Panosu", kalin=True, boyut=14, yazi=KOYU_LACIVERT)
    kpiler = [
        (1, "Toplam Hasılat (TL)", "=HESAP!B6", TL),
        (2, "Toplam Reçete Maliyeti (TL)", "=HESAP!B7", TL),
        (3, "TOPLAM KAYIP/KAÇAK (TL)", "=HESAP!B11", TL),
        (4, "Toplam Kayıp Oranı", "=HESAP!B12", YÜZDE),
        (5, "Kâr Katkısı (TL)", "=HESAP!B8", TL),
        (6, "Kâr Katkısı Kayıp Oranı", "=HESAP!B20", YÜZDE),
        (7, "Karar", "=KARAR!B11", NORMAL),
        (8, "Kalite Skoru", "=KONTROLLER!B16", CATI),
    ]
    for kolon, ad, form, sayi in kpiler:
        h(ws, 3, kolon, ad, hiza="center", kalin=True, yazi="FFFFFF",
          zemin=KOYU_LACIVERT, boyut=10, kaydir=True)
        h(ws, 4, kolon, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True, hiza="center")
    h(ws, 6, 1, "Karar", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 4, "En Kritik Hammadde", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    h(ws, 6, 5, "=HESAP!B16", boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 6, 7, "En Yüksek Kayıp Tutarı (TL)", kalin=True, boyut=11, yazi=KOYU_LACIVERT, kaydir=True)
    h(ws, 6, 8, "=HESAP!B17", sayi=TL, boyut=12, kalin=True, yazi=DIS_REF_YEŞIL)
    ws.row_dimensions[6].height = 30
    # Ürün bazlı hasılat/maliyet/kâr
    h(ws, 10, 1, "Ürün", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 2, "Hasılat (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 3, "Reçete Maliyeti (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 4, "Kâr Katkısı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 10, 5, "Kâr Katkısı Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    urun_etiketleri = ["UR-01 Klasik Pide", "UR-02 Lahmacun", "UR-03 Kumpir",
                       "UR-04 Tost Sandviç", "UR-05 Çiğ Köfte Menü", "UR-06 İskender"]
    for i, etiket in enumerate(urun_etiketleri, 11):
        h(ws, i, 1, etiket, yazi=GRİ)
        h(ws, i, 2, f"=HESAP!B{i+13}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f"=HESAP!D{i+13}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 4, f"=HESAP!F{i+13}", sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 5, f"=IFERROR(HESAP!F{i+13}/HESAP!B{i+13},0)", sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g1 = BarChart()
    g1.type = "col"
    g1.title = "Ürün Bazlı Hasılat ve Reçete Maliyeti"
    g1.add_data(Reference(ws, min_col=2, max_col=3, min_row=10, max_row=17), titles_from_data=True)
    g1.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g1.height, g1.width = 9, 15
    ws.add_chart(g1, "F8")
    g2 = BarChart()
    g2.type = "col"
    g2.title = "Ürün Bazlı Kâr Katkısı"
    g2.add_data(Reference(ws, min_col=4, min_row=10, max_row=17), titles_from_data=True)
    g2.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g2.height, g2.width = 9, 12
    ws.add_chart(g2, "U8")
    g7 = BarChart()
    g7.type = "col"
    g7.title = "Ürün Bazlı Kâr Katkısı Oranı"
    g7.add_data(Reference(ws, min_col=5, min_row=10, max_row=17), titles_from_data=True)
    g7.set_categories(Reference(ws, min_col=1, min_row=11, max_row=17))
    g7.height, g7.width = 9, 12
    ws.add_chart(g7, "AH8")
    # Hammadde kayıp dağılımı
    h(ws, 19, 1, "Hammadde", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 2, "Kayıp Tutarı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 19, 3, "Kayıp Oranı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ham in enumerate(["Un", "Tereyağı", "Peynir", "Yumurta", "Et", "Domates", "Salça", "Patates"], 20):
        h(ws, i, 1, ham, yazi=GRİ)
        h(ws, i, 2, f'=SUMIFS(tblStok[Kayıp Tutarı (TL)],tblStok[Hammadde],$A{i})',
          sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, i, 3, f'=IFERROR(SUMIFS(tblStok[Kayıp/Kaçak (Miktar)],tblStok[Hammadde],$A{i})/SUMIFS(tblStok[Teorik Tüketim],tblStok[Hammadde],$A{i}),0)',
          sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
    g3 = PieChart()
    g3.title = "Hammadde Bazlı Kayıp Tutarı Payı"
    g3.add_data(Reference(ws, min_col=2, min_row=19, max_row=27), titles_from_data=True)
    g3.set_categories(Reference(ws, min_col=1, min_row=20, max_row=27))
    g3.height, g3.width = 9, 12
    ws.add_chart(g3, "F21")
    g8 = BarChart()
    g8.type = "col"
    g8.title = "Hammadde Bazlı Kayıp Oranı"
    g8.add_data(Reference(ws, min_col=3, min_row=19, max_row=27), titles_from_data=True)
    g8.set_categories(Reference(ws, min_col=1, min_row=20, max_row=27))
    g8.height, g8.width = 9, 12
    ws.add_chart(g8, "AH21")
    # Kayıp trendi
    h(ws, 29, 1, "Dönem", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 29, 2, "Kayıp Tutarı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i in range(4):
        h(ws, 30 + i, 1, i + 1, sayi=CATI, yazi=GRİ)
        h(ws, 30 + i, 2, f"=HESAP!B{33+i}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g4 = LineChart()
    g4.title = "Kayıp Tutarı Trendi"
    g4.add_data(Reference(ws, min_col=2, min_row=29, max_row=34), titles_from_data=True)
    g4.set_categories(Reference(ws, min_col=1, min_row=30, max_row=34))
    g4.height, g4.width = 9, 15
    ws.add_chart(g4, "F34")
    # Projeksiyon
    h(ws, 36, 1, "Ortalama Kayıp (TL)", yazi="333333")
    h(ws, 36, 2, "=HESAP!B38", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 37, 1, "Gelecek Dönem Tahmini (TL)", kalin=True, yazi=KOYU_LACIVERT)
    h(ws, 37, 2, "=HESAP!B39", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 38, 1, "Alt Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 38, 2, "=HESAP!B41", sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 39, 1, "Üst Tahmin Sınırı (TL)", yazi="333333")
    h(ws, 39, 2, "=HESAP!B42", sayi=TL, yazi=DIS_REF_YEŞIL)
    g5 = BarChart()
    g5.type = "col"
    g5.title = "Projeksiyon: Ortalama, Tahmin ve Bant"
    g5.add_data(Reference(ws, min_col=2, min_row=36, max_row=39), titles_from_data=True)
    g5.set_categories(Reference(ws, min_col=1, min_row=37, max_row=39))
    g5.height, g5.width = 9, 15
    ws.add_chart(g5, "U40")
    # Senaryo verisi
    h(ws, 41, 5, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 41, 6, "Kayıp Tutarı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    for i, ad in enumerate(["İyimser", "Baz", "Kötümser", "Kritik"], 42):
        h(ws, i, 5, ad, yazi=GRİ)
        h(ws, i, 6, f"=SENARYO_DUYARLILIK!C{i-35}", sayi=TL, yazi=DIS_REF_YEŞIL)
    g6 = LineChart()
    g6.title = "Senaryo Kayıp Tutarları"
    g6.add_data(Reference(ws, min_col=6, min_row=41, max_row=46), titles_from_data=True)
    g6.set_categories(Reference(ws, min_col=5, min_row=42, max_row=46))
    g6.height, g6.width = 9, 12
    ws.add_chart(g6, "F48")
    baski_hazirla(ws, "A1:M48", URUN_AD)
    alt_bant(ws, 48, "Paneldeki tüm göstergeler canlı formüllerden beslenir.")
    genislik(ws, {"A": 30, "B": 20, "C": 22, "D": 18, "E": 20, "F": 20, "G": 20, "H": 20})


def senaryo_duyarlilik(ws):
    sayfa_hazirla(ws, "SENARYO_DUYARLILIK", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Senaryo ve Duyarlılık Analizi", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Kayıp oranı çarpanına göre senaryo bant aralığı ve hammadde duyarlılık "
                "(tornado) analizi.", yazi=GRİ, kaydir=True)
    ws.row_dimensions[4].height = 26
    h(ws, 6, 1, "Senaryo", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 2, "Kayıp Çarpanı", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 6, 3, "Kayıp Tutarı (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    senaryolar = [
        (7, "İyimser", "=IF(COUNTA(tblStok[Hammadde])>0,MutfakSenIyiKayip,0)",
         "=IFERROR(HESAP!B11*MutfakSenIyiKayip/MutfakSenBazKayip,0)"),
        (8, "Baz", "=IF(COUNTA(tblStok[Hammadde])>0,MutfakSenBazKayip,0)", "=HESAP!B11"),
        (9, "Kötümser", "=IF(COUNTA(tblStok[Hammadde])>0,MutfakSenKotuKayip,0)",
         "=IFERROR(HESAP!B11*MutfakSenKotuKayip/MutfakSenBazKayip,0)"),
        (10, "Kritik", "=IF(COUNTA(tblStok[Hammadde])>0,MutfakSenKritikKayip,0)",
         "=IFERROR(HESAP!B11*MutfakSenKritikKayip/MutfakSenBazKayip,0)"),
    ]
    for satir, ad, form, c in senaryolar:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=YÜZDE, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 3, c, sayi=TL, yazi=DIS_REF_YEŞIL)
    h(ws, 12, 1, "Tornado Analizi", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 13, 1, "Değişken", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 13, 2, "Düşük Senaryo (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    h(ws, 13, 3, "Yüksek Senaryo (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI, kaydir=True)
    h(ws, 13, 4, "Etki Genişliği (TL)", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    tornado = [
        (14, "Et Birim Maliyeti", "=HESAP!B11*MutfakTornadoEtOran", "=HESAP!B11*MutfakTornadoEtOran*3"),
        (15, "Fire Oranı", "=HESAP!B11*MutfakTornadoFireOran", "=HESAP!B11*MutfakTornadoFireOran*3"),
    ]
    for satir, ad, dus, yuk in tornado:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, dus, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 3, yuk, sayi=TL, yazi=DIS_REF_YEŞIL)
        h(ws, satir, 4, f'=C{satir}-B{satir}', sayi=TL, yazi=DIS_REF_YEŞIL)
    baski_hazirla(ws, "A1:D10", URUN_AD)
    alt_bant(ws, 18, "Tornado genişliği hangi değişkenin kayıp tutarını en çok etkilediğini gösterir.")
    genislik(ws, {"A": 28, "B": 24, "C": 24, "D": 22})


def rapor(ws):
    sayfa_hazirla(ws, "RAPOR", RENK, URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Yönetici Raporu", kalin=True, boyut=15, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "=CONCATENATE(\"Rapor Tarihi: \",TEXT(MutfakRaporTarihi,\"gg.aa.yyyy\"),\" | \",MutfakFirmaUnvan)",
      yazi=GRİ)
    h(ws, 6, 1, "Karar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 6, 2, "=KARAR!B11", boyut=13, kalin=True, yazi=DIS_REF_YEŞIL)
    h(ws, 7, 1, "Gerekçe", yazi="333333", kaydir=True)
    h(ws, 7, 2, "=KARAR!B12", yazi=DIS_REF_YEŞIL, kaydir=True)
    ws.row_dimensions[7].height = 40
    h(ws, 9, 1, "Ana Göstergeler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    rapor_satirlari = [
        (10, "Toplam Hasılat (TL)", "=HESAP!B6", TL),
        (11, "Toplam Reçete Maliyeti (TL)", "=HESAP!B7", TL),
        (12, "Toplam Kâr Katkısı (TL)", "=HESAP!B8", TL),
        (13, "Toplam Kayıp/Kaçak (TL)", "=HESAP!B11", TL),
        (14, "Toplam Kayıp Oranı", "=HESAP!B12", YÜZDE),
        (15, "Kâr Katkısı Kayıp Oranı", "=HESAP!B20", YÜZDE),
        (16, "En Kritik Hammadde", "=HESAP!B16", NORMAL),
    ]
    for satir, ad, form, sayi in rapor_satirlari:
        h(ws, satir, 1, ad, yazi="333333")
        h(ws, satir, 2, form, sayi=sayi, yazi=DIS_REF_YEŞIL, kalin=True)
    h(ws, 18, 1, "Varsayımlar ve Uyarılar", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 19, 1, "Kayıp oranı eşikleri, senaryo ve tornado çarpanları AYARLAR'dan okunur; "
                 "güncellenebilir.", yazi=GRİ, kaydir=True)
    ws.merge_cells("A19:P19")
    h(ws, 20, 1, "Kayıp tespiti fiili sayımla doğrulanmalıdır; bu dosya karar destek aracıdır.",
      yazi=GRİ, kaydir=True)
    ws.merge_cells("A20:P20")
    h(ws, 22, 1, "=CONCATENATE(\"Sürüm \",MutfakDosyaSurumu,\" | \",MutfakRaporHazirlayan)",
      yazi=GRİ, boyut=9)
    ws.merge_cells("A22:P22")
    baski_hazirla(ws, "A1:P22", URUN_AD)
    genislik(ws, {"A": 40, "B": 26})


def ornek_veri(ws):
    sayfa_hazirla(ws, "ORNEK_VERI", RENK, URUN_AD, son_kolon=36)
    h(ws, 3, 1, "Örnek Veriler", kalin=True, boyut=13, yazi=KOYU_LACIVERT, kaydir=True)
    ws.row_dimensions[3].height = 30
    h(ws, 4, 1, "Ürünü tanımak için bu sayfadaki örnek stok, reçete ve ürün verileri "
                "yan yana bloklar halinde sunulur; ana sayfalarda denemeler yapın.",
      yazi=GRİ, kaydir=True)
    # ---- Blok 1: Örnek Stok (A-L) ----
    h(ws, 5, 1, "Örnek Stok", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar = ["Hammadde", "Birim", "Dönem Başı Stok", "Alım", "Dönem Sonu Stok",
                "Fiili Tüketim", "Birim Maliyet (TL)", "Teorik Tüketim",
                "Kayıp/Kaçak (Miktar)", "Kayıp Oranı", "Kayıp Tutarı (TL)", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar, basla=1)
    formuller = {
        "Fiili Tüketim": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",MAX(0,tblOrnekStok[[#This Row],[Dönem Başı Stok]]+tblOrnekStok[[#This Row],[Alım]]-tblOrnekStok[[#This Row],[Dönem Sonu Stok]]))',
        "Teorik Tüketim": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",SUMIFS(tblOrnekRecete[Teorik Tüketim],tblOrnekRecete[Hammadde],tblOrnekStok[[#This Row],[Hammadde]]))',
        "Kayıp/Kaçak (Miktar)": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",MAX(0,tblOrnekStok[[#This Row],[Fiili Tüketim]]-tblOrnekStok[[#This Row],[Teorik Tüketim]]))',
        "Kayıp Oranı": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",IF(tblOrnekStok[[#This Row],[Teorik Tüketim]]=0,0,tblOrnekStok[[#This Row],[Kayıp/Kaçak (Miktar)]]/tblOrnekStok[[#This Row],[Teorik Tüketim]]))',
        "Kayıp Tutarı (TL)": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",tblOrnekStok[[#This Row],[Kayıp/Kaçak (Miktar)]]*tblOrnekStok[[#This Row],[Birim Maliyet (TL)]])',
        "Açıklama": '=IF(tblOrnekStok[[#This Row],[Hammadde]]="","",CONCATENATE("Hammadde: ",tblOrnekStok[[#This Row],[Hammadde]]," | Kayıp: ",TEXT(tblOrnekStok[[#This Row],[Kayıp Tutarı (TL)]],"₺ #,##0")))',
    }
    tablo_ekle(ws, "tblOrnekStok", "A6:L1005", sutunlar, formuller)
    ornek_stok = [
        ("Un", "kg", 120, 90, 135, 18),
        ("Tereyağı", "kg", 18, 14, 12, 120),
        ("Peynir", "kg", 30, 25, 20, 160),
        ("Yumurta", "adet", 400, 300, 350, 6),
        ("Et", "kg", 60, 55, 40, 320),
        ("Domates", "kg", 45, 40, 30, 35),
        ("Salça", "kg", 12, 10, 8, 60),
        ("Patates", "kg", 100, 80, 70, 22),
    ]
    for i, (ham, birim, bas, alim, son, maliyet) in enumerate(ornek_stok, 7):
        for kolon, deger in [(1, ham), (2, birim), (3, bas), (4, alim), (5, son), (7, maliyet)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=3).number_format = CATI
        for kolon in (4, 5, 6, 8, 9):
            ws.cell(row=satir, column=kolon).number_format = CATI
        ws.cell(row=satir, column=7).number_format = TL
        ws.cell(row=satir, column=10).number_format = YÜZDE
        ws.cell(row=satir, column=11).number_format = TL
    giris_hucreleri(ws, 7, 1005, [1, 2, 3, 4, 5, 7])
    # ---- Blok 2: Örnek Reçeteler (N-W) ----
    h(ws, 5, 14, "Örnek Reçeteler", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar2 = ["Ürün Kodu", "Ürün Adı", "Hammadde", "Birim", "Miktar",
                 "Birim Maliyet (TL)", "Tutar (TL)", "Satış Adedi (Adet)",
                 "Teorik Tüketim", "Açıklama"]
    baslik_satiri(ws, 6, sutunlar2, basla=14)
    formuller2 = {
        "Tutar (TL)": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="","",tblOrnekRecete[[#This Row],[Miktar]]*tblOrnekRecete[[#This Row],[Birim Maliyet (TL)]])',
        "Satış Adedi (Adet)": '=IF(tblOrnekRecete[[#This Row],[Ürün Kodu]]="","",IFERROR(INDEX(tblOrnekUrun[Satış Adedi (Adet)],MATCH(tblOrnekRecete[[#This Row],[Ürün Kodu]],tblOrnekUrun[Ürün Kodu],0)),0))',
        "Teorik Tüketim": '=IF(tblOrnekRecete[[#This Row],[Miktar]]="","",tblOrnekRecete[[#This Row],[Miktar]]*tblOrnekRecete[[#This Row],[Satış Adedi (Adet)]])',
        "Açıklama": '=IF(tblOrnekRecete[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblOrnekRecete[[#This Row],[Ürün Adı]]))',
    }
    tablo_ekle(ws, "tblOrnekRecete", "N6:W1005", sutunlar2, formuller2)
    ornek_recete = [
        ("UR-01", "Klasik Pide", "Un", "kg", 0.250, 18),
        ("UR-01", "Klasik Pide", "Tereyağı", "kg", 0.030, 120),
        ("UR-01", "Klasik Pide", "Peynir", "kg", 0.080, 160),
        ("UR-01", "Klasik Pide", "Yumurta", "adet", 1, 6),
        ("UR-02", "Lahmacun", "Un", "kg", 0.080, 18),
        ("UR-02", "Lahmacun", "Et", "kg", 0.040, 320),
        ("UR-02", "Lahmacun", "Domates", "kg", 0.050, 35),
        ("UR-02", "Lahmacun", "Salça", "kg", 0.010, 60),
        ("UR-03", "Kumpir", "Patates", "kg", 0.350, 22),
        ("UR-03", "Kumpir", "Tereyağı", "kg", 0.040, 120),
        ("UR-03", "Kumpir", "Peynir", "kg", 0.050, 160),
        ("UR-04", "Tost Sandviç", "Un", "kg", 0.120, 18),
        ("UR-04", "Tost Sandviç", "Peynir", "kg", 0.060, 160),
        ("UR-04", "Tost Sandviç", "Tereyağı", "kg", 0.020, 120),
        ("UR-05", "Çiğ Köfte Menü", "Un", "kg", 0.100, 18),
        ("UR-05", "Çiğ Köfte Menü", "Domates", "kg", 0.040, 35),
        ("UR-05", "Çiğ Köfte Menü", "Salça", "kg", 0.020, 60),
        ("UR-06", "İskender", "Et", "kg", 0.180, 320),
        ("UR-06", "İskender", "Tereyağı", "kg", 0.030, 120),
        ("UR-06", "İskender", "Patates", "kg", 0.150, 22),
    ]
    for i, (kod, urun, ham, birim, miktar, maliyet) in enumerate(ornek_recete, 7):
        for kolon, deger in [(14, kod), (15, urun), (16, ham), (17, birim), (18, miktar), (19, maliyet)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=18).number_format = CATI
        for kolon in (19, 20):
            ws.cell(row=satir, column=kolon).number_format = TL
        ws.cell(row=satir, column=21).number_format = CATI
        ws.cell(row=satir, column=22).number_format = CATI
    giris_hucreleri(ws, 7, 1005, [14, 15, 16, 17, 18, 19])
    # ---- Blok 3: Örnek Ürünler (Y-AF) ----
    h(ws, 5, 25, "Örnek Ürünler", kalin=True, boyut=11, yazi=KOYU_LACIVERT)
    sutunlar3 = ["Ürün Kodu", "Ürün Adı", "Satış Adedi (Adet)", "Birim Satış Fiyatı (TL)",
                 "Reçete Maliyeti (TL)", "Satış Hasılatı (TL)", "Kâr Katkısı (TL)",
                 "Açıklama"]
    baslik_satiri(ws, 6, sutunlar3, basla=25)
    formuller3 = {
        "Reçete Maliyeti (TL)": '=IF(tblOrnekUrun[[#This Row],[Ürün Kodu]]="","",SUMIFS(tblOrnekRecete[Tutar (TL)],tblOrnekRecete[Ürün Kodu],tblOrnekUrun[[#This Row],[Ürün Kodu]]))',
        "Satış Hasılatı (TL)": '=IF(tblOrnekUrun[[#This Row],[Ürün Kodu]]="","",tblOrnekUrun[[#This Row],[Satış Adedi (Adet)]]*tblOrnekUrun[[#This Row],[Birim Satış Fiyatı (TL)]])',
        "Kâr Katkısı (TL)": '=IF(tblOrnekUrun[[#This Row],[Ürün Kodu]]="","",tblOrnekUrun[[#This Row],[Satış Hasılatı (TL)]]-tblOrnekUrun[[#This Row],[Reçete Maliyeti (TL)]])',
        "Açıklama": '=IF(tblOrnekUrun[[#This Row],[Ürün Adı]]="","",CONCATENATE("Ürün: ",tblOrnekUrun[[#This Row],[Ürün Adı]]))',
    }
    tablo_ekle(ws, "tblOrnekUrun", "Y6:AF1005", sutunlar3, formuller3)
    ornek_urun = [
        ("UR-01", "Klasik Pide", 320, 120),
        ("UR-02", "Lahmacun", 410, 95),
        ("UR-03", "Kumpir", 280, 140),
        ("UR-04", "Tost Sandviç", 520, 75),
        ("UR-05", "Çiğ Köfte Menü", 190, 150),
        ("UR-06", "İskender", 230, 180),
    ]
    for i, (kod, urun, adet, fiyat) in enumerate(ornek_urun, 7):
        for kolon, deger in [(25, kod), (26, urun), (27, adet), (28, fiyat)]:
            ws.cell(row=i, column=kolon).value = deger
    for satir in range(7, 1005):
        ws.cell(row=satir, column=27).number_format = CATI
        for kolon in (28, 29, 30, 31):
            ws.cell(row=satir, column=kolon).number_format = TL
    giris_hucreleri(ws, 7, 1005, [25, 26, 27, 28])
    sabitle(ws, "A7")
    genislik(ws, {"A": 14, "B": 8, "C": 15, "D": 9, "E": 15, "F": 13, "G": 15, "H": 14,
                  "I": 17, "J": 11, "K": 15, "L": 34,
                  "N": 10, "O": 15, "P": 14, "Q": 8, "R": 9, "S": 14, "T": 12, "U": 14,
                  "V": 13, "W": 26,
                  "Y": 10, "Z": 15, "AA": 14, "AB": 16, "AC": 16, "AD": 15, "AE": 13,
                  "AF": 24})


def degisiklik_kaydi(ws):
    sayfa_hazirla(ws, "DEGISIKLIK_KAYDI", "8D8D8D", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Değişiklik Kaydı", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    baslik_satiri(ws, 4, ["Tarih", "Sürüm", "Değişiklik", "Neden", "Yapan"])
    kayit = ["10.08.2026", "1.0.0", "İlk üretim", "Ürünün yayına hazırlanması", "ExcelArşiv"]
    for i, deger in enumerate(kayit, 1):
        h(ws, 5, i, deger, yazi="333333")
    genislik(ws, {"A": 16, "B": 12, "C": 40, "D": 32, "E": 18})


def listeler(ws):
    sayfa_hazirla(ws, "LISTELER", "808080", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Bağlı Listeler", kalin=True, boyut=12, yazi=KOYU_LACIVERT)
    h(ws, 4, 1, "Bu sayfadaki listeler açılır menüleri besler; yeni seçenek eklemek için "
                "sarı hücrelere yazın.", yazi=GRİ, kaydir=True)
    h(ws, 6, 1, "Hammaddeler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, ham in enumerate(HAMMADDELER, 7):
        h(ws, i, 1, ham, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"A{i}", "Tanım: Hammadde | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni hammadde için satır ekleyin")
    h(ws, 7, 3, "Birimler", kalin=True, yazi="FFFFFF", zemin=ORTA_MAVI)
    for i, birim in enumerate(BIRIMLER, 7):
        h(ws, i, 3, birim, zemin=GIRIS_SARI)
        yorum_ekle(ws, f"C{i}", "Tanım: Birim | Neden önemli: Açılır listeyi besler | "
                                 "Doğru kullanım: Yeni birim için satır ekleyin")
    genislik(ws, {"A": 22, "C": 16})


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
        (7, "MutfakFirmaUnvan", "Örnek Lokanta A.Ş.", "Metin", None, None, "Firma unvanı rapor başlığında görünür.", "Kullanıcı"),
        (8, "MutfakRaporTarihi", TARIH_D, "Tarih", None, None, "Raporun dayanak tarihi.", "Kullanıcı"),
        (9, "MutfakRaporHazirlayan", "Mutfak ve Mali İşler Yöneticisi", "Metin", None, None, "Raporu hazırlayan kişi.", "Kullanıcı"),
        (10, "MutfakDosyaSurumu", SURUM, "Metin", None, None, "Dosya sürümü.", "Kullanıcı"),
        (11, "MutfakInceleEsik", 0.03, "Oran", 0, 1, "Toplam kayıp oranının inceleme eşiği (%3).", "Varsayım"),
        (12, "MutfakKritikEsik", 0.08, "Oran", 0, 1, "Toplam kayıp oranının kritik eşiği (%8).", "Varsayım"),
        (13, "MutfakKaliteIyi", 90, "Puan", 0, 100, "Yüksek kalite skoru eşiği.", "Varsayım"),
        (14, "MutfakKaliteOrta", 70, "Puan", 0, 100, "Orta kalite skoru eşiği.", "Varsayım"),
        (15, "MutfakSenIyiKayip", 0.50, "Oran", 0, 3, "İyimser senaryo kayıp çarpanı.", "Varsayım"),
        (16, "MutfakSenBazKayip", 1.00, "Oran", 0, 3, "Baz senaryo kayıp çarpanı.", "Varsayım"),
        (17, "MutfakSenKotuKayip", 1.50, "Oran", 0, 3, "Kötümser senaryo kayıp çarpanı.", "Varsayım"),
        (18, "MutfakSenKritikKayip", 2.00, "Oran", 0, 3, "Kritik senaryo kayıp çarpanı.", "Varsayım"),
        (19, "MutfakTahminCarpan", 1.645, "Çarpan", 0, 10, "Tahmin güven bandı çarpanı (%90 normal dağılım).", "Varsayım"),
        (20, "MutfakTahminDonem", 5, "Dönem", 1, 12, "Gelecek dönem tahmin noktası.", "Varsayım"),
        (21, "MutfakTornadoEtOran", 0.10, "Oran", 0, 0.5, "Et birim maliyeti tornado adımı (%10).", "Varsayım"),
        (22, "MutfakTornadoFireOran", 0.05, "Oran", 0, 0.5, "Fire/kayıp tornado adımı.", "Varsayım"),
        (23, "MutfakP90Oran", 0.90, "Oran", 0, 1, "Birim maliyet dağılımında üst yüzdelik dilimin oranı.", "Varsayım"),
        (24, "MutfakYuksekKayipEsik", 50000, "TL", 0, 1000000000, "Yüksek kayıp tutarı kontrol eşiği.", "Varsayım"),
        (25, "MutfakAksiyon1_ACIKLAMA",
         "Toplam kayıp oranı kritik eşiğin üzerinde; kayıp yoğunlaşan hammaddelerde sayım, tartım ve fire prosedürünü denetleyin.", "Metin", None, None,
         "KRİTİK KAYIP kararında gösterilen ilk öneri.", "Varsayım"),
        (26, "MutfakAksiyon2_ACIKLAMA",
         "Kayıp oranı inceleme aralığında; yüksek kayıplı kalemlerde alım, depolama ve porsiyon kontrolünü gözden geçirin.", "Metin", None, None,
         "İNCELE kararında gösterilen ikinci öneri.", "Varsayım"),
        (27, "MutfakAksiyon3_ACIKLAMA",
         "Kayıp kontrol altında; dönemlik mutabakatı sürdürerek kayıp eğilimini izlemeye devam edin.", "Metin", None, None,
         "KONTROL ALTINDA kararında gösterilen üçüncü öneri.", "Varsayım"),
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
        h(ws, satir, 1, f"MutfakPanelSatir_{i}", kalin=True, zemin=ACIK_GRI, boyut=9)
        hucre = h(ws, satir, 2, i, zemin=GIRIS_SARI)
        hucre.number_format = CATI
        h(ws, satir, 3, "Satır", boyut=9, yazi=GRİ)
        h(ws, satir, 6, f"PANO'daki grafiğin ürün sırası ({i}. veri).", boyut=9, yazi=GRİ, kaydir=True)
        h(ws, satir, 7, "Kullanıcı", boyut=9, yazi=GRİ)
        h(ws, satir, 8, YURURLUK, boyut=9, yazi=GRİ, sayi=TARİH)
        h(ws, satir, 9, DOGRULAMA, boyut=9, yazi=GRİ, sayi=TARİH)
        yorum_ekle(ws, f"B{satir}",
                   f"Tanım: Panel sıra numarası {i} | "
                   f"Neden önemli: Grafikleri besler | "
                   f"Doğru kullanım: Ürün sırasını girin")
        ws.row_dimensions[satir].height = 26
    genislik(ws, {"A": 34, "B": 24, "C": 12, "D": 10, "E": 10, "F": 60, "G": 22, "H": 16, "I": 18})


def kilavuz(ws):
    sayfa_hazirla(ws, "KILAVUZ", "333F50", URUN_AD, son_kolon=30)
    h(ws, 3, 1, "Kullanım Kılavuzu", kalin=True, boyut=13, yazi=KOYU_LACIVERT)
    bolumler = [
        ("Veri girişi",
         "URUN_LISTESI sayfasında ürün kodu/adı, satış adedi ve birim satış fiyatını; RECETE "
         "sayfasında her ürünün hammadde ihtiyacını; STOK sayfasında dönem başı, alım ve dönem "
         "sonu stoklarını sarı hücrelere girin. Tutarlar otomatik hesaplanır."),
        ("Teorik ve fiili tüketim",
         "Teorik tüketim, satış adedinin reçetedeki hammadde miktarıyla çarpılmasıyla ürün "
         "bazında toplanır. Fiili tüketim ise dönem başı + alım − dönem sonu stoktur. Kayıp, "
         "fiili − teorik farkın pozitif kısmıdır."),
        ("Kayıp tutarı ve oranı",
         "Kayıp miktarı birim maliyetle çarpılarak kayıp tutarı bulunur; kayıp oranı kaybın "
         "teorik tüketime bölümüdür. Toplam kayıp, kâr katkısından düşülerek net kâr katkısı üretilir."),
        ("Karar ve aksiyon",
         "KARAR sayfası toplam kayıp oranını eşiklerle karşılaştırır; KONTROL ALTINDA / DÜŞÜK "
         "KAYIP / İNCELE / KRİTİK KAYIP kararını gerekçesiyle üretir ve önerilen aksiyonları gösterir."),
        ("Oranlar ve senaryolar",
         "AYARLAR sayfasından inceleme/kritik eşikleri, senaryo ve tornado çarpanlarını "
         "güncelleyin; PANO ve SENARYO_DUYARLILIK anında güncellenir."),
        ("Güven ve karar kalitesi",
         "Veri kalite skoru girişlerin tamlığını ölçer. Stok verisi yokken karar VERİ YOK; "
         "negatif girişte DURDUR üretir. Verileriniz cihazınızdan çıkmaz; dosya tamamen çevrimdışıdır."),
        ("Uyumluluk",
         "Excel 2016-365 (Windows ve Mac), LibreOffice ve Google Sheets ile uyumludur. "
         "Makro içermez; formüller ayrıştırıcı ile denetlenmiştir."),
        ("Karar kuralları",
         "KARAR sayfasındaki sonuç şu kurallarla üretilir: (1) stok verisi yoksa VERİ YOK; "
         "(2) negatif stok girişi varsa DURDUR; (3) kayıp oranı kritik eşiğin üzerindeyse KRİTİK "
         "KAYIP; (4) oran inceleme aralığındaysa İNCELE; (5) oran sıfırın üzerinde ama eşiğin "
         "altındaysa DÜŞÜK KAYIP; (6) aksi durumda KONTROL ALTINDA. Gerekçe, kararın altında "
         "makine tarafından üretilir."),
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
                                  ("KRİTİK KAYIP", kirmizi, kirmizi_font), ("İNCELE", sari, None),
                                  ("DÜŞÜK KAYIP", sari, None), ("KONTROL ALTINDA", yesil, yesil_font)):
        ws.conditional_formatting.add("B11",
            CellIsRule(operator="equal", formula=[f'"{deger}"'], fill=dolu, font=yazifont))

    ws = wb["PANO"]
    for r in (4, 6):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"VERİ YOK"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DURDUR"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"KRİTİK KAYIP"'], fill=kirmizi, font=kirmizi_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"İNCELE"'], fill=sari))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"DÜŞÜK KAYIP"'], fill=sari))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"KONTROL ALTINDA"'], fill=yesil, font=yesil_font))

    ws = wb["STOK"]
    for c in ("C", "D", "E"):
        ws.conditional_formatting.add(f"{c}6:{c}1004",
            CellIsRule(operator="lessThan", formula=["0"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("J6:J1004",
        CellIsRule(operator="greaterThan", formula=["MutfakInceleEsik"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("J6:J1004",
        CellIsRule(operator="greaterThan", formula=["MutfakKritikEsik"], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("K6:K1004",
        CellIsRule(operator="greaterThan", formula=["0"], fill=sari, font=kirmizi_font))

    ws = wb["HESAP"]
    ws.conditional_formatting.add("B11",
        CellIsRule(operator="greaterThan", formula=["MutfakYuksekKayipEsik"],
                   fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B12",
        CellIsRule(operator="greaterThan", formula=["MutfakInceleEsik"], fill=sari, font=kirmizi_font))
    ws.conditional_formatting.add("B12",
        CellIsRule(operator="greaterThan", formula=["MutfakKritikEsik"], fill=kirmizi, font=kirmizi_font))

    ws = wb["KONTROLLER"]
    for r in range(6, 11):
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="equal", formula=['"TEMİZ"'], fill=yesil, font=yesil_font))
        ws.conditional_formatting.add(f"B{r}",
            CellIsRule(operator="notEqual", formula=['"TEMİZ"'], fill=kirmizi, font=kirmizi_font))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="greaterThanOrEqual", formula=["90"], fill=yesil, font=yesil_font))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="between", formula=["70", "89"], fill=sari))
    ws.conditional_formatting.add("B16",
        CellIsRule(operator="lessThan", formula=["70"], fill=kirmizi, font=kirmizi_font))


def moduller(wb):
    # Analitik modül değer/yorumları KONTROLLER sayfasında hücre olarak durur;
    # ad tanımları bu hücrelere işaret eder (D03/D04/D05 gereği).
    ws = wb["KONTROLLER"]
    modul_satirlari = {
        "modulT1OrtKayip": ("T1", "Ortalama Kayıp Oranı",
                            "=IFERROR(AVERAGE(tblStok[Kayıp Oranı]),0)",
                            '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama hammadde kayıp oranı ",TEXT(C24,"0.0%")," — mutfaktaki genel kayıp düzeyini gösterir")',
                            YÜZDE),
        "modulT2OrtMaliyet": ("T2", "Ortalama Reçete Maliyeti",
                              "=IFERROR(AVERAGE(tblUrun[Reçete Maliyeti (TL)]),0)",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Ortalama ürün reçete maliyeti ",TEXT(C25,"₺ #,##0")," TL — menü kârlılık düzeyini gösterir")',
                              TL),
        "modulO1KayipSapmasi": ("O1", "Kayıp Tutarı Sapması (MAD)",
                                "=IFERROR(AVEDEV(tblStok[Kayıp Tutarı (TL)]),0)",
                                '=_xlfn.TEXTJOIN("; ",TRUE,"Hammadde kayıp tutarı ortalama mutlak sapması ",TEXT(C26,"₺ #,##0")," TL — kaybın kalemler arası heterojenliğini ölçer")',
                                TL),
        "modulO2EnYuksekKayip": ("O2", "En Yüksek Kayıp Tutarı",
                                 "=IFERROR(MAX(tblStok[Kayıp Tutarı (TL)]),0)",
                                 '=_xlfn.TEXTJOIN("; ",TRUE,"En yüksek hammadde kayıp tutarı ",TEXT(C27,"₺ #,##0")," TL — öncelikli müdahale kalemidir")',
                                 TL),
        "modulO3Senaryo": ("O3", "Senaryo Bant Genişliği",
                           "=IFERROR(STDEV.P(SENARYO_DUYARLILIK!C7:C10),0)",
                           '=_xlfn.TEXTJOIN("; ",TRUE,"Senaryo kayıp tutarı standart sapması ",TEXT(C28,"₺ #,##0")," TL — kayıp belirsizliğinin etkisini gösterir")',
                           TL),
        "modulO6Hhi": ("O6", "Kayıp Yoğunlaşması (HHI)",
                       '=IFERROR(IF(SUM(tblStok[Kayıp Tutarı (TL)])=0,0,SUMPRODUCT((tblStok[Kayıp Tutarı (TL)]/SUM(tblStok[Kayıp Tutarı (TL)]))^2)),0)',
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Hammadde kayıp yoğunlaşma endeksi ",TEXT(C29,"0.000")," — 1 e yaklaştıkça kayıp tek kalemde toplanır")',
                       None),
        "modulO8KaliteSkor": ("O8", "Veri Kalite Skoru",
                              "=KONTROLLER!B16",
                              '=_xlfn.TEXTJOIN("; ",TRUE,"Veri kalite skoru ",TEXT(C30,"0")," puan — ",IF(C30>=MutfakKaliteIyi,"yüksek güven","veri girişi tamamlanmalı"))',
                              CATI),
        "modulI1Tahmin": ("I1", "Gelecek Dönem Kayıp Tahmini",
                          "=HESAP!B39",
                          '=_xlfn.TEXTJOIN("; ",TRUE,"Gelecek dönem kayıp tahmini ",TEXT(C31,"₺ #,##0")," TL — bant aralığı için alt/üst sınır HESAP sayfasındadır")',
                          TL),
        "modulI2P90": ("I2", "Birim Maliyet Yüzdelik P90",
                       "=IFERROR(_xlfn.PERCENTILE.INC(tblStok[Birim Maliyet (TL)],MutfakP90Oran),0)",
                       '=_xlfn.TEXTJOIN("; ",TRUE,"Birim maliyetlerin %90 ı ",TEXT(C32,"₺ #,##0")," TL altındadır — maliyet dağılımının üst sınırını gösterir")',
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
    sayfa_sirasi = ["KAPAK", "HIZLI_BASLANGIC", "URUN_LISTESI", "RECETE", "STOK", "HESAP",
                    "KONTROLLER", "KARAR", "PANO", "SENARYO_DUYARLILIK", "RAPOR", "ORNEK_VERI",
                    "DEGISIKLIK_KAYDI", "LISTELER", "AYARLAR", "KILAVUZ"]
    for ad in sayfa_sirasi:
        wb.create_sheet(ad)
    ilk = wb["Sheet"]
    wb.remove(ilk)

    kapak(wb["KAPAK"])
    hizli_baslangic(wb["HIZLI_BASLANGIC"])
    urun_listesi(wb["URUN_LISTESI"])
    recete(wb["RECETE"])
    stok(wb["STOK"])
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
        if isinstance(anahtar, str) and anahtar.startswith("Mutfak"):
            ad_ekle(wb, anahtar, f"AYARLAR!$B${satir}")

    # Kritik çıktı ad tanımları (D02 karşılıkları ve PANO beslemeleri)
    ad_ekle(wb, "MutfakToplamKayip", "HESAP!$B$11")
    ad_ekle(wb, "MutfakToplamKayipOrani", "HESAP!$B$12")
    ad_ekle(wb, "MutfakKarar", "KARAR!$B$11")
    ad_ekle(wb, "MutfakEnKritikHammadde", "HESAP!$B$16")
    ad_ekle(wb, "MutfakProjX", "HESAP!$A$33:$A$36")

    ad_ekle(wb, "ListeHammaddeler", "LISTELER!$A$7:$A$17")
    ad_ekle(wb, "ListeBirimler", "LISTELER!$C$7:$C$12")

    for ws in wb.worksheets:
        sayfa_koru(ws)

    dosya = "MutfakKayipKacakHesaplayici.xlsx"
    wb.save(dosya)
    print("Dosya oluşturuldu:", dosya)


if __name__ == "__main__":
    main()
