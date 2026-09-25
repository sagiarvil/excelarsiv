import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.colors import ColorChoice

wb_path = 'EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx'
wb = openpyxl.load_workbook(wb_path)
ws = wb['PANO']

# ------------------------------------------------------------------
# GRAFİK 1: DÖNEMSEL İCRA VE PROJEKSİYON (ÇİFT SÜTUN BAR GRAFİĞİ)
# Kaynak Veri: B14:E20 (B14..E14 başlıklar, B15..B20 dönemler, C15..C20 Aktif Hacim, D15..D20 Maliyet, E15..E20 Hedef)
# Konum: B28:G44 (Hemen alt kısma ferahça yerleşim)
# ------------------------------------------------------------------
chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "DÖNEMSEL BORÇ HACMİ VE MALİYET TRENDİ"
chart1.y_axis.title = "Endeks / Milyon TL"
chart1.x_axis.title = "Dönemler"
chart1.width = 16.5
chart1.height = 9.5

data1 = Reference(ws, min_col=3, min_row=14, max_col=5, max_row=20)
cats1 = Reference(ws, min_col=2, min_row=15, max_row=20)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)

# Renkler & Tasarım
ws.add_chart(chart1, "B28")

# ------------------------------------------------------------------
# GRAFİK 2: KREDİ PORTFÖYÜ BANKA / SEGMENT DAĞILIMI (HALKA / DOUGHNUT GRAFİK)
# Kaynak Veri: H14:I18 (H14: Parametre/Kalem, I14: Ana Tutar)
# Konum: H28:M44 (Sağ tarafa simetrik yerleşim)
# ------------------------------------------------------------------
chart2 = DoughnutChart()
chart2.title = "BORÇ DAĞILIMI VE RİSK YOĞUNLAŞMASI"
chart2.style = 10
chart2.width = 14.5
chart2.height = 9.5
chart2.holeSize = 60

data2 = Reference(ws, min_col=9, min_row=14, max_row=18)
cats2 = Reference(ws, min_col=8, min_row=15, max_row=18)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)

ws.add_chart(chart2, "H28")

wb.save(wb_path)
print("PANO sayfasına 2 adet $1,000 değerinde yönetici grafik yapısı (Dönemsel Bar + Portföy Halka) başarıyla eklendi.")
