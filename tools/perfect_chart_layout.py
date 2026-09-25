import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, Reference, Series
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList

wb_path = 'EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx'
wb = openpyxl.load_workbook(wb_path)
ws = wb['PANO']

# Mevcut grafikleri kaldırıp sıfırdan kusursuz ölçülerle ekleyelim
ws._charts.clear()

# 1. Grafiklerin oturacağı satırların (28..46) yüksekliklerini 18.0 pt olarak sabitleyelim
# 18 satır x 18pt = 324pt (~11.4 cm) tuval alanı
for r in range(28, 48):
    ws.row_dimensions[r].height = 18.0

# -------------------------------------------------------------
# GRAFİK 1: DÖNEMSEL BORÇ HACMİ VE MALİYET TRENDİ (ÇİFT SÜTUN BAR)
# Kolonlar: B..G (Genişlik: 76.0 birim -> ~15.5 cm)
# -------------------------------------------------------------
chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "DÖNEMSEL BORÇ HACMİ VE MALİYET TRENDİ"
chart1.y_axis.title = "Tutar (Milyon TL)"
chart1.x_axis.title = "Dönemler"
chart1.width = 15.5
chart1.height = 10.5

# Başlık ve eksenlerin yazı boyutları & netlik
chart1.legend.legendPos = "t" # Üstte ferah gösterge (ezilme/baskılama yok)

data1 = Reference(ws, min_col=3, min_row=14, max_col=5, max_row=20)
cats1 = Reference(ws, min_col=2, min_row=15, max_row=20)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)

# Renkler
colors1 = ['1B365D', 'D97706', '059669'] # Lacivert, Amber, Zümrüt
for idx, s in enumerate(chart1.series):
    s.graphicalProperties.solidFill = colors1[idx % len(colors1)]

ws.add_chart(chart1, "B28")

# -------------------------------------------------------------
# GRAFİK 2: BORÇ DAĞILIMI VE RİSK YOĞUNLAŞMASI (MODERN HALKA)
# Kolonlar: H..M (Genişlik: 108.0 birim -> ~20.5 cm)
# -------------------------------------------------------------
chart2 = DoughnutChart()
chart2.title = "KONSOLİDE RİSK VE BORÇ DAĞILIMI"
chart2.style = 10
chart2.width = 20.0
chart2.height = 10.5
chart2.holeSize = 60

# Göstergeyi sağa (r) alarak dilimlerin ve rakamların birbirini ezmesini engelle
chart2.legend.legendPos = "r"

data2 = Reference(ws, min_col=9, min_row=14, max_row=18)
cats2 = Reference(ws, min_col=8, min_row=15, max_row=18)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)

# Dilim renkleri (İnfografik uyumu)
slice_colors = ['1B365D', '2563EB', 'D97706', '059669']
for idx, col in enumerate(slice_colors):
    dp = DataPoint(idx=idx)
    dp.graphicalProperties.solidFill = col
    chart2.series[0].data_points.append(dp)

# Rakamların net okunması için Data Labels ekle (değerler dilimlerin üzerinde görünür)
chart2.dataLabels = DataLabelList()
chart2.dataLabels.showVal = False
chart2.dataLabels.showPercent = True # Yüzdeyi net göster

ws.add_chart(chart2, "H28")

# -------------------------------------------------------------
# 3. VERİ SAYFALARINDAKİ SÜTUN GENİŞLİKLERİ VE KESİNTİSİZLİK BUFFERI
# -------------------------------------------------------------
# Kolonlardaki sayısal ve metinsel verilerin asla "###" veya yarım kesilmeye uğramaması için
# her sütunun genişliği en az `max_len + 4.5` olarak genişletilir.
for sname in ['BANKALAR', 'LIMITLER', 'KREDILER', 'TAKSIT_PLANI', 'ODEME_TAKVIMI', 'MASRAF_KOMISYON', 'TEMINATLAR', 'YENI_TEKLIFLER', 'REFINANSMAN']:
    ws_d = wb[sname]
    for c in range(1, 15):
        col_let = openpyxl.utils.get_column_letter(c)
        cur_w = ws_d.column_dimensions[col_let].width or 13.0
        # Hücreleri tara
        max_len = 0
        for r in range(5, min(ws_d.max_row + 1, 50)):
            v = ws_d.cell(r, c).value
            if v is not None and not str(v).startswith('='):
                max_len = max(max_len, len(str(v)))
        header = ws_d.cell(5, c).value
        if header:
            max_len = max(max_len, len(str(header)))
            needed_w = max(cur_w, max_len + 4.5)
            ws_d.column_dimensions[col_let].width = round(needed_w, 1)

wb.save(wb_path)
print("PANO grafikleri ve veri sütunları sıfır ezilme, sıfır sürtünme ve tam netlik ile yeniden dizayn edildi.")
