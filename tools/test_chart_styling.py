import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, Reference, Series
from openpyxl.drawing.colors import ColorChoice

wb = openpyxl.load_workbook('EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx')
ws = wb['PANO']

ch0 = ws._charts[0]
colors = ['1B365D', 'D97706', '059669'] # Navy, Amber, Emerald
for idx, s in enumerate(ch0.series):
    s.graphicalProperties.solidFill = colors[idx % len(colors)]

ch1 = ws._charts[1]
# Doughnut slices colors
slice_colors = ['1B365D', '2563EB', 'D97706', '059669']
from openpyxl.chart.series import DataPoint
for idx, col in enumerate(slice_colors):
    dp = DataPoint(idx=idx)
    dp.graphicalProperties.solidFill = col
    ch1.series[0].data_points.append(dp)

wb.save('EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx')
print("Grafik renkleri ve dilim stilleri başarıyla uygulandı.")
