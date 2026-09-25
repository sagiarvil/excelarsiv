import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb_path = 'EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx'
wb = openpyxl.load_workbook(wb_path)

# -------------------------------------------------------------
# 1. PANO GELİŞTİRMESİ: SÜTUN VE KART BORDER SİMETRİSİ (SIEMENS GRADE)
# -------------------------------------------------------------
ws_pano = wb['PANO']
ws_pano.views.sheetView[0].showGridLines = False

# H14:L19 tablosu için sütun genişliklerini altın oranda ayarla (sıfır kırılma)
# H: PARAMETRE / KALEM (30), I: ANA TUTAR (16), J: FARK / MARJ (16), K: VADE (10), L: KONTROL DURUMU (26)
ws_pano.column_dimensions['H'].width = 30.0
ws_pano.column_dimensions['I'].width = 16.0
ws_pano.column_dimensions['J'].width = 16.0
ws_pano.column_dimensions['K'].width = 10.0
ws_pano.column_dimensions['L'].width = 26.0

# KPI Kartlarındaki hücre içi dikey border artıklarını temizle
# Kartlar: B..D, E..G, H..J, K..M
cards = [
    ('B', 'C', 'D'),
    ('E', 'F', 'G'),
    ('H', 'I', 'J'),
    ('K', 'L', 'M')
]

thin_gray = Side(style='thin', color='E2E8F0')
med_navy = Side(style='medium', color='1E3A8A')
no_border = Side(style=None)

for c_left, c_mid, c_right in cards:
    for r in range(5, 10):
        # Sol hücre
        left_side = med_navy if c_left in ['B', 'E', 'H', 'K'] else thin_gray
        top_side = thin_gray if r == 5 else no_border
        bottom_side = thin_gray if r == 9 else no_border
        
        ws_pano[f'{c_left}{r}'].border = Border(top=top_side, bottom=bottom_side, left=left_side, right=no_border)
        ws_pano[f'{c_mid}{r}'].border = Border(top=top_side, bottom=bottom_side, left=no_border, right=no_border)
        ws_pano[f'{c_right}{r}'].border = Border(top=top_side, bottom=bottom_side, left=no_border, right=thin_gray)

# -------------------------------------------------------------
# 2. YÖNETİM RAPORU: KUTU VE SÜTUN SİMETRİSİ
# -------------------------------------------------------------
ws_rap = wb['YÖNETİM_RAPORU']
ws_rap.views.sheetView[0].showGridLines = False

# Kart B..F ve H..L
box1_cols = ['B', 'C', 'D', 'E', 'F']
box2_cols = ['H', 'I', 'J', 'K', 'L']

for r in range(6, 12):
    top_side = thin_gray if r == 6 else no_border
    bottom_side = thin_gray if r == 11 else no_border
    
    # Kutu 1
    for c in box1_cols:
        l_side = thin_gray if c == 'B' else no_border
        r_side = thin_gray if c == 'F' else no_border
        ws_rap[f'{c}{r}'].border = Border(top=top_side, bottom=bottom_side, left=l_side, right=r_side)
        
    # Kutu 2
    for c in box2_cols:
        l_side = thin_gray if c == 'H' else no_border
        r_side = thin_gray if c == 'L' else no_border
        ws_rap[f'{c}{r}'].border = Border(top=top_side, bottom=bottom_side, left=l_side, right=r_side)

# -------------------------------------------------------------
# 3. YENİ SAYFA: 00_HAKKINDA_VE_REHBER (5 YAŞINDA ÇOCUĞA ANLATIR GİBİ SİSTEM MİMARİSİ)
# -------------------------------------------------------------
guide_sheet_name = '00_HAKKINDA_VE_REHBER'
if guide_sheet_name in wb.sheetnames:
    del wb[guide_sheet_name]

ws_info = wb.create_sheet(title=guide_sheet_name, index=1) # KILAVUZ'un hemen ardından 2. sıraya
ws_info.views.sheetView[0].showGridLines = False
ws_info.sheet_properties.tabColor = '002563EB' # Siemens / Royal Blue

# Sütun Genişlikleri
ws_info.column_dimensions['A'].width = 3.5
ws_info.column_dimensions['B'].width = 24.0
ws_info.column_dimensions['C'].width = 32.0
ws_info.column_dimensions['D'].width = 40.0
ws_info.column_dimensions['E'].width = 24.0
ws_info.column_dimensions['F'].width = 18.0

# Başlık Satırları
ws_info.row_dimensions[2].height = 36.0
ws_info['B2'] = "SİSTEM HAKKINDA, KULLANIM REHBERİ VE TERİMLER SÖZLÜĞÜ"
ws_info['B2'].font = Font(name='Montserrat', size=18, bold=True, color='0F172A')

ws_info.row_dimensions[3].height = 20.0
ws_info['B3'] = "5 Yaşındaki Birine Anlatır Gibi Basit, Sade ve Eksiksiz Finansal Karar Mimarisi · Siemens & Enterprise Standartları"
ws_info['B3'].font = Font(name='Segoe UI', size=10, italic=True, color='475569')

# Bölüm 1: 3 ALTIN SORUDA BU SİSTEM NEDİR?
ws_info.row_dimensions[5].height = 24.0
ws_info['B5'] = "1. TEMEL MANTIK — 3 DAKİKADA SİSTEMİN ÖZÜ"
ws_info['B5'].font = Font(name='Montserrat', size=12, bold=True, color='1E3A8A')

questions = [
    ("SORU", "GÜNLÜK DİLLE CEVAP (5 YAŞINDA ÇOCUĞA ANLATIR GİBİ)", "ŞİRKETİNİZE FAYDASI NE?", "İLGİLİ SAYFA"),
    ("1. Bu dosya ne işe yarar?", 
     "Tıpkı arabanın gösterge paneli gibidir. Hangi bankaya ne kadar borcunuz olduğunu, bu ay ne kadar taksit ödeyeceğinizi ve paranızın yetip yetmeyeceğini tek bir ekranda gösterir.", 
     "Gözden kaçan, sürpriz veya geciken taksitleri sıfırlar; faiz cezası ve itibar kaybını önler.", 
     "PANO, YÖNETİM_RAPORU"),
    ("2. Kullanıcı olarak ben ne yapacağım?", 
     "Çok kolay! Yalnızca 'GİRDİLER' sekmesindeki sarı renkli kutulara çektiğiniz kredileri ve vadesini yazacaksınız. Geri kalan tüm hesapları sistem kendi yapar.", 
     "Veri girişi birkaç dakikada biter, karmaşık formüllerle veya Excel kodlarıyla uğraşmazsınız.", 
     "GİRDİLER, KREDILER"),
    ("3. Sistem bana ne kazandırır?", 
     "Hangi bankaya çok faiz ödediğinizi bulur. 'Şu krediyi kapatıp şuraya taşırsan ayda 495.000 TL tasarruf edersin' diyerek şirketinize somut nakit kazandırır.", 
     "Gereksiz faiz ve masrafları yok eder; banka pazarlıklarında şirketinizi güçlü kılar.", 
     "REFINANSMAN, KARAR_MOTORU")
]

ws_info.row_dimensions[7].height = 24.0
ws_info['B7'] = questions[0][0]
ws_info['C7'] = questions[0][1]
ws_info['D7'] = questions[0][2]
ws_info['E7'] = questions[0][3]

navy_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
white_bold = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')

for col_let in ['B', 'C', 'D', 'E']:
    cell = ws_info[f'{col_let}7']
    cell.fill = navy_fill
    cell.font = white_bold
    cell.alignment = Alignment(horizontal='center' if col_let in ['B', 'E'] else 'left', vertical='center')

for idx, q in enumerate(questions[1:], start=8):
    ws_info.row_dimensions[idx].height = 42.0
    ws_info[f'B{idx}'] = q[0]
    ws_info[f'C{idx}'] = q[1]
    ws_info[f'D{idx}'] = q[2]
    ws_info[f'E{idx}'] = q[3]
    
    bg_color = 'F8FAFC' if idx % 2 == 0 else 'FFFFFF'
    row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    
    ws_info[f'B{idx}'].font = Font(name='Segoe UI', size=9.5, bold=True, color='0F172A')
    ws_info[f'C{idx}'].font = Font(name='Segoe UI', size=9, color='334155')
    ws_info[f'D{idx}'].font = Font(name='Segoe UI', size=9, color='059669', bold=True)
    ws_info[f'E{idx}'].font = Font(name='Segoe UI', size=9, color='2563EB', bold=True)
    
    for col_let in ['B', 'C', 'D', 'E']:
        c = ws_info[f'{col_let}{idx}']
        c.fill = row_fill
        c.alignment = Alignment(horizontal='center' if col_let in ['B', 'E'] else 'left', vertical='center', wrap_text=True)
        c.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

# Bölüm 2: TERİMLER SÖZLÜĞÜ (EN KARMAŞIK FİNANSAL TERİMLERİN SADECE AÇIKLAMASI)
ws_info.row_dimensions[13].height = 24.0
ws_info['B13'] = "2. SÖZLÜK — BU KELİMELER NE ANLAMA GELİYOR?"
ws_info['B13'].font = Font(name='Montserrat', size=12, bold=True, color='1E3A8A')

terms = [
    ("TERİM / KAVRAM", "ÇOK BASİT AÇIKLAMASI", "BU ŞABLONDA NEREDE HESAPLANIR?", "ÖRNEK / İPUCU"),
    ("Refinansman", "Pahalı olan eski krediyi kapatıp yerine daha düşük faizli yeni kredi çekmek. Yani borcu ucuza taşımak.", "REFINANSMAN sayfası", "Ör: %45 faizli krediyi kapatıp %38'e çekerek ayda 495.714 TL kâr etmek."),
    ("Borç Servisi", "Gelecek 30 gün içinde bankalara tıkır tıkır ödemeniz gereken anapara ve faiz taksitlerinin toplamı.", "PANO (E7 hücresi)", "Mevcut 30 günlük borç servisimiz: ₺934.864"),
    ("Stres Testi", "'Gelecek ay işler çok kötü giderse ve kasaya %20 daha az para girerse borcumuzu ödeyebilir miyiz?' provası.", "YÖNETİM_RAPORU", "Şirketimiz -%20 krizde bile temerrüt riski yaşamadan borcunu ödeyebilmektedir."),
    ("Temerrüt Riski", "Taksit gününde kasada para bulunamaması ve borcun icralık / takipli hale gelme tehlikesi.", "KARAR_MOTORU", "Sistem yeşil 'UYGUN' yakıyorsa temerrüt riskiniz %0 demektir."),
    ("DSCR (Borç Karşılama)", "Kasaya giren serbest paranın taksitlere olan oranı. 1.0 üzerindeyse şirket rahat nefes alır.", "KARAR_MOTORU (B14)", "1.25 üstü = Mükemmel Güvenlik Bölgesi."),
    ("Amortisman Takvimi", "Kredinin ilk gününden son gününe kadar her ay kaç para faiz, kaç para anapara ödeneceğini gösteren liste.", "TAKSIT_PLANI", "Toplam kaç ay taksit kaldığını tek tıkla listeler.")
]

ws_info.row_dimensions[15].height = 24.0
ws_info['B15'] = terms[0][0]
ws_info['C15'] = terms[0][1]
ws_info['D15'] = terms[0][2]
ws_info['E15'] = terms[0][3]

for col_let in ['B', 'C', 'D', 'E']:
    cell = ws_info[f'{col_let}15']
    cell.fill = navy_fill
    cell.font = white_bold
    cell.alignment = Alignment(horizontal='center' if col_let in ['B', 'C'] else 'left', vertical='center')

for idx, t in enumerate(terms[1:], start=16):
    ws_info.row_dimensions[idx].height = 36.0
    ws_info[f'B{idx}'] = t[0]
    ws_info[f'C{idx}'] = t[1]
    ws_info[f'D{idx}'] = t[2]
    ws_info[f'E{idx}'] = t[3]
    
    bg_color = 'F8FAFC' if idx % 2 == 0 else 'FFFFFF'
    row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    
    ws_info[f'B{idx}'].font = Font(name='Segoe UI', size=9.5, bold=True, color='0F172A')
    ws_info[f'C{idx}'].font = Font(name='Segoe UI', size=9, color='334155')
    ws_info[f'D{idx}'].font = Font(name='Segoe UI', size=9, color='2563EB', bold=True)
    ws_info[f'E{idx}'].font = Font(name='Segoe UI', size=8.5, italic=True, color='64748B')
    
    for col_let in ['B', 'C', 'D', 'E']:
        c = ws_info[f'{col_let}{idx}']
        c.fill = row_fill
        c.alignment = Alignment(horizontal='center' if col_let == 'B' else 'left', vertical='center', wrap_text=True)
        c.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

# Bölüm 3: SİSTEMİN 5 AŞAMALI NAVİGASYON ŞEMASI
ws_info.row_dimensions[24].height = 24.0
ws_info['B24'] = "3. SİSTEMİN 5 AŞAMALI AKIŞ ŞEMASI"
ws_info['B24'].font = Font(name='Montserrat', size=12, bold=True, color='1E3A8A')

stages = [
    ("AŞAMA", "GÖREVİ", "KULLANILACAK SEKME", "GİRİLECEK VERİ", "ÜRETİLEN SONUÇ"),
    ("1. Hazırlık", "Şirket parametrelerini ve nakit durumunu tanıtma", "AYARLAR, LISTELER", "Firma adı, serbest nakit miktarı", "Tüm formüllerin eşik değerleri hazır olur"),
    ("2. Veri Girişi", "Bankalardaki borçları ve kredi takvimini yazma", "GİRDİLER, KREDILER, LIMITLER", "Kredi tutarı, faiz, taksit tarihi", "Ödeme takvimi otomatik üretilir"),
    ("3. Anlık Takip", "Borç yükünü ve 30 günlük ödemeleri izleme", "PANO", "Otomatik hesaplanır", "Gözden kaçan taksit riski kalmaz"),
    ("4. Fırsat Avı", "Pahalı kredileri ucuza taşıma & Kriz testi", "REFINANSMAN, YÖNETİM_RAPORU", "Yeni teklif faizleri", "Aylık ₺495.714'ye varan nakit tasarrufu"),
    ("5. Sunum & Karar", "Yönetim kuruluna ve bankalara resmi sunum", "RAPOR", "Tek tıkla hazır", "A4 dikey resmi finansal rapor")
]

ws_info.row_dimensions[26].height = 24.0
ws_info['B26'] = stages[0][0]
ws_info['C26'] = stages[0][1]
ws_info['D26'] = stages[0][2]
ws_info['E26'] = stages[0][3]
ws_info['F26'] = stages[0][4]

for col_let in ['B', 'C', 'D', 'E', 'F']:
    cell = ws_info[f'{col_let}26']
    cell.fill = navy_fill
    cell.font = white_bold
    cell.alignment = Alignment(horizontal='center' if col_let in ['B', 'D'] else 'left', vertical='center')

for idx, s in enumerate(stages[1:], start=27):
    ws_info.row_dimensions[idx].height = 32.0
    ws_info[f'B{idx}'] = s[0]
    ws_info[f'C{idx}'] = s[1]
    ws_info[f'D{idx}'] = s[2]
    ws_info[f'E{idx}'] = s[3]
    ws_info[f'F{idx}'] = s[4]
    
    bg_color = 'F8FAFC' if idx % 2 == 0 else 'FFFFFF'
    row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    
    ws_info[f'B{idx}'].font = Font(name='Segoe UI', size=9.5, bold=True, color='1D4ED8')
    ws_info[f'C{idx}'].font = Font(name='Segoe UI', size=9, bold=True, color='0F172A')
    ws_info[f'D{idx}'].font = Font(name='Segoe UI', size=9, color='2563EB')
    ws_info[f'E{idx}'].font = Font(name='Segoe UI', size=9, color='475569')
    ws_info[f'F{idx}'].font = Font(name='Segoe UI', size=9, bold=True, color='059669')
    
    for col_let in ['B', 'C', 'D', 'E', 'F']:
        c = ws_info[f'{col_let}{idx}']
        c.fill = row_fill
        c.alignment = Alignment(horizontal='center' if col_let in ['B', 'D'] else 'left', vertical='center')
        c.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

# KILAVUZ sayfasına bu yeni rehbere doğrudan giden buton ekle
ws_k = wb['KILAVUZ']
ws_k['I7'] = "KULLANIM REHBERİ & SÖZLÜK"
ws_k['I9'] = "5 yaşındaki birine anlatır gibi sadeleştirilmiş sistem özeti, kullanım kılavuzu ve finans terimleri sözlüğü için tıklayınız."
ws_k['I14'] = "📖 REHBERİ VE SÖZLÜĞÜ AÇ (#00_HAKKINDA_VE_REHBER!A1)"

wb.save(wb_path)
print("Siemens Standartlarında Kurumsal Mimari & 00_HAKKINDA_VE_REHBER Sekmesi Başarıyla Eklendi.")
