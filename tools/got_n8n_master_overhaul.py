import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, Reference, Series
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

prod_path = 'EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx'
wb = openpyxl.load_workbook(prod_path)

# ==============================================================================
# BÖLÜM 1: KILAVUZ SAYFASINDAKİ TABLO SÜTUN KIRILMALARININ SIFIRLANMASI
# ==============================================================================
ws_k = wb['KILAVUZ']
ws_k.views.sheetView[0].showGridLines = False

# 21. satırdan 27. satıra kadar olan TÜM merged hücreleri güvenli kaldır
ranges_to_remove = []
for mr in ws_k.merged_cells.ranges:
    for r in range(21, 28):
        if str(r) in str(mr):
            ranges_to_remove.append(mr)
            break

for mr in ranges_to_remove:
    ws_k.merged_cells.remove(mr)

# Kolon Genişlikleri:
# B: AŞAMA NO (Genişlik: 12)
# C: AŞAMA ADI (Genişlik: 28)
# D..M: İŞLEM DETAYI VE MANTIK (D22:M22 genişletilmiş merge)
# N..O: KULLANILAN SAYFALAR (N22:O22 merge)
# P..R: DOĞRUDAN ERİŞİM BUTONU (P22:R22 merge)
ws_k.column_dimensions['B'].width = 12.0
ws_k.column_dimensions['C'].width = 30.0

navy_fill = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
white_bold = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
thin_border = Border(top=Side(style='thin', color='CBD5E1'),
                     bottom=Side(style='thin', color='CBD5E1'),
                     left=Side(style='thin', color='CBD5E1'),
                     right=Side(style='thin', color='CBD5E1'))

# Başlık Satırı 21
ws_k.row_dimensions[21].height = 26.0
ws_k['B21'] = "AŞAMA NO"
ws_k['C21'] = "AŞAMA ADI & KAPSAM"
ws_k['D21'] = "İŞLEM DETAYI VE YALIN MANTIK"
ws_k.merge_cells('D21:M21')
ws_k['N21'] = "İLGİLİ SAYFALAR"
ws_k.merge_cells('N21:O21')
ws_k['P21'] = "HIZLI ERİŞİM"
ws_k.merge_cells('P21:R21')

for col in ['B', 'C', 'D', 'N', 'P']:
    cell = ws_k[f'{col}21']
    cell.fill = navy_fill
    cell.font = white_bold
    cell.alignment = Alignment(horizontal='center' if col in ['B', 'N', 'P'] else 'left', vertical='center', wrap_text=False)

# Tablo Veri Satırları 22..26
stages_data = [
    ("AŞAMA 1", "SİSTEM KURULUMU & PARAMETRELER", "Şirket unvanı, raporlama tarihi, serbest nakit ve risk toleranslarını tanımlayın.", "AYARLAR, LISTELER", "⚙ AYARLAR (#AYARLAR!A1)"),
    ("AŞAMA 2", "STRATEJİK KOKPİT & İZLEME", "Kredi portföy büyüklüğü, 30 günlük ödeme baskısı ve likidite radarını tek ekrandan izleyin.", "PANO", "▶ PANO KOKPİTİ (#PANO!A1)"),
    ("AŞAMA 3", "PORTFÖY VE TAKSİT GİRİŞİ", "Banka limitlerini, çekilen kredileri ve sözleşme amortisman takvimlerini ilgili sayfalara girin.", "GİRDİLER, KREDİLER", "📋 GİRİŞLERİ DÜZENLE (#GİRDİLER!A1)"),
    ("AŞAMA 4", "REFİNANSMAN VE STRES TESTİ", "Krediyi başka bankaya taşıma faiz kârını hesaplayın. -%20 nakit daralması kriz testini inceleyin.", "REFİNANSMAN, RAPOR", "📊 STRES TESTİ (#YÖNETİM_RAPORU!A1)"),
    ("AŞAMA 5", "YÖNETİM KURULU VE BANKA ÇIKTISI", "Kredi komiteleri ve yönetim kuruluna sunulacak A4 dikey baskıya hazır kurumsal rapor çıktısını alın.", "RAPOR", "📄 A4 RAPORU (#RAPOR!A1)")
]

for idx, (asama, ad, detay, sayfalar, buton) in enumerate(stages_data, start=22):
    ws_k.row_dimensions[idx].height = 36.0
    
    # B: Aşama No
    ws_k[f'B{idx}'] = asama
    ws_k[f'B{idx}'].font = Font(name='Segoe UI', size=9, bold=True, color='1D4ED8')
    ws_k[f'B{idx}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
    
    # C: Aşama Adı
    ws_k[f'C{idx}'] = ad
    ws_k[f'C{idx}'].font = Font(name='Segoe UI', size=9, bold=True, color='0F172A')
    ws_k[f'C{idx}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
    
    # D: Detay (Merge D..M)
    ws_k[f'D{idx}'] = detay
    ws_k[f'D{idx}'].font = Font(name='Segoe UI', size=8.5, color='475569')
    ws_k[f'D{idx}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_k.merge_cells(f'D{idx}:M{idx}')
    
    # N: Sayfalar (Merge N..O)
    ws_k[f'N{idx}'] = sayfalar
    ws_k[f'N{idx}'].font = Font(name='Segoe UI', size=9, bold=True, color='059669')
    ws_k[f'N{idx}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
    ws_k.merge_cells(f'N{idx}:O{idx}')
    
    # P: Buton (Merge P..R)
    ws_k[f'P{idx}'] = buton
    ws_k[f'P{idx}'].font = Font(name='Segoe UI', size=8.5, bold=True, color='FFFFFF')
    ws_k[f'P{idx}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
    ws_k[f'P{idx}'].fill = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
    ws_k.merge_cells(f'P{idx}:R{idx}')
    
    # Border
    for c in range(2, 19):
        col_let = openpyxl.utils.get_column_letter(c)
        ws_k[f'{col_let}{idx}'].border = thin_border

# ==============================================================================
# BÖLÜM 2: PANO GRAFİKLERİNİN VE METİNLERİNİN KUSURSUZ İZOLASYONU
# ==============================================================================
ws_p = wb['PANO']
ws_p.views.sheetView[0].showGridLines = False

# Grafikleri temizle ve 30. satırdan başlat (27..29 tampon boşluk)
ws_p._charts.clear()

ws_p.row_dimensions[27].height = 14.0
ws_p.row_dimensions[28].height = 14.0
ws_p.row_dimensions[29].height = 14.0

# 30 ila 52. satırları boş tuval olarak 18pt'ye sabitle
for r in range(30, 53):
    ws_p.row_dimensions[r].height = 18.0
    for c in range(1, 15):
        cell = ws_p.cell(r, c)
        cell.value = None
        cell.fill = PatternFill(fill_type=None)
        cell.border = Border()

# GRAFİK 1: DÖNEMSEL BORÇ HACMİ VE MALİYET TRENDİ (ÇİFT SÜTUN BAR)
chart1 = BarChart()
chart1.type = "col"
chart1.style = 10
chart1.title = "DÖNEMSEL BORÇ HACMİ VE MALİYET TRENDİ"
chart1.y_axis.title = "Tutar (Milyon TL)"
chart1.x_axis.title = "Dönemler"
chart1.width = 16.0
chart1.height = 10.0
# Üst üste binmeyi engellemek için Lejant'ı sağa (r) alıyoruz!
chart1.legend.legendPos = "r"

data1 = Reference(ws_p, min_col=3, min_row=14, max_col=5, max_row=20)
cats1 = Reference(ws_p, min_col=2, min_row=15, max_row=20)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)

colors1 = ['1B365D', 'D97706', '059669']
for idx, s in enumerate(chart1.series):
    s.graphicalProperties.solidFill = colors1[idx % len(colors1)]

ws_p.add_chart(chart1, "B30")

# GRAFİK 2: BORÇ DAĞILIMI VE RİSK YOĞUNLAŞMASI (MODERN HALKA)
chart2 = DoughnutChart()
chart2.title = "KONSOLİDE BORÇ DAĞILIMI"
chart2.style = 10
chart2.width = 18.0
chart2.height = 10.0
chart2.holeSize = 60
chart2.legend.legendPos = "r"

data2 = Reference(ws_p, min_col=9, min_row=14, max_row=18)
cats2 = Reference(ws_p, min_col=8, min_row=15, max_row=18)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)

slice_colors = ['1B365D', '2563EB', 'D97706', '059669']
for idx, col in enumerate(slice_colors):
    dp = DataPoint(idx=idx)
    dp.graphicalProperties.solidFill = col
    chart2.series[0].data_points.append(dp)

chart2.dataLabels = None

ws_p.add_chart(chart2, "H30")

wb.save(prod_path)
print("GoT & n8n Tamamlandı: KILAVUZ tablosu ve PANO grafikleri kusursuzlaştırıldı.")
