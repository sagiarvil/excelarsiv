import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

wb_path = Path("/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx")
wb = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=False)

# Fonts
FONT_DISPLAY = "Montserrat"
FONT_BODY = "Segoe UI"

# Colors
C_WHITE = "FFFFFF"
C_NAVY_DARK = "0F172A"       # Slate 900
C_NAVY_HEADER = "1E293B"     # Slate 800
C_CARD_BG = "F8FAFC"         # Slate 50
C_ZEBRA = "F1F5F9"           # Slate 100
C_BORDER = "CBD5E1"          # Slate 300
C_BORDER_LIGHT = "E2E8F0"

# Text colors
C_TEXT_MAIN = "0F172A"
C_TEXT_SUB = "334155"
C_TEXT_MUTED = "64748B"

# Accents
C_EMERALD_BG = "ECFDF5"
C_EMERALD_TXT = "047857"
C_EMERALD_LINE = "10B981"
C_EMERALD_DARK = "065F46"

C_BLUE_BG = "EFF6FF"
C_BLUE_TXT = "1D4ED8"
C_BLUE_LINE = "3B82F6"
C_BLUE_LIGHT = "60A5FA"

C_RED_BG = "FEF2F2"
C_RED_TXT = "B91C1C"
C_RED_LINE = "EF4444"

C_PURPLE_BG = "FAF5FF"
C_PURPLE_TXT = "6B21A8"
C_PURPLE_LINE = "A855F7"

C_INPUT_YELLOW = "FFFBEB"
C_INPUT_BLUE = "1D4ED8"

# Borders
border_thin = Border(
    left=Side(style='thin', color=C_BORDER), right=Side(style='thin', color=C_BORDER),
    top=Side(style='thin', color=C_BORDER), bottom=Side(style='thin', color=C_BORDER)
)
border_header = Border(
    left=Side(style='thin', color="334155"), right=Side(style='thin', color="334155"),
    top=Side(style='thin', color="334155"), bottom=Side(style='medium', color="0F172A")
)

print("1. Re-ordering sheets according to the 5-Stage CFO Process Workflow...")

# 5-STAGE PROCESS WORKFLOW ORDER
process_order = [
    "KILAVUZ",             # Aşama 1: Giriş, Tanıtım & 5 Aşamalı Proses Haritası
    "PANO",                # Aşama 2: Stratejik Kokpit & Likidite Erken Uyarı Radarı
    "GİRDİLER",            # Aşama 3: Hızlı İşlem & Günlük Ödeme Takip Havuzu
    "BANKALAR",            # Aşama 3.1: Banka İlişkileri ve Şube Yönetimi
    "LIMITLER",            # Aşama 3.2: Kredi Limit Doluluk ve Boşluk Analizi
    "KREDILER",            # Aşama 3.3: Kredi Sözleşme ve Portföy Havuzu
    "TAKSIT_PLANI",        # Aşama 3.4: Dönem Bazında Amortisman Planı
    "ODEME_TAKVIMI",       # Aşama 3.5: 30 / 60 / 90 Gün Nakit Akış Takvimi
    "MASRAF_KOMISYON",     # Aşama 3.6: Gizli Banka Masrafları & Komisyonlar
    "TEMINATLAR",          # Aşama 3.7: İpotek, Teminat Mektubu & Kefaletler
    "YENI_TEKLIFLER",      # Aşama 3.8: Banka Kredi ve Refinansman Teklif Havuzu
    "REFINANSMAN",         # Aşama 4.1: Borç Yapılandırma & Arbitraj Analizi
    "YÖNETİM_RAPORU",      # Aşama 4.2: C-Level Stres Testi (-%20 Şok Senaryosu)
    "RAPOR",               # Aşama 4.3: Bankaya/Yönetime Sunulacak A4 PDF Karnesi
    "KARAR_MOTORU",        # Aşama 5.1: Otomatik Karar & Aksiyon Algoritması
    "MOTOR",               # Aşama 5.2: Analitik Modüller Ara Hesaplama Motoru
    "SENARYO_DUYARLILIK",  # Aşama 5.3: Faiz/Kur Duyarlılık & Tornado Motoru
    "KONTROLLER",          # Aşama 5.4: Veri Bütünlüğü & Denge Kontrolleri
    "AYARLAR",             # Aşama 5.5: Risk Eşikleri & Parametre Havuzu
    "LISTELER",            # Aşama 5.6: Dropdown Doğrulama Veri Tabanı
    # Master Governance Layer
    "00_BASLANGIC", "01_YURUTME_CEKIRDEGI", "02_KONTROL_LISTESI", "03_GORSEL_TASARIM", "04_MUHASEBE_THP", "05_KARAR_SATIS",
    # Supporting technical sheets
    "KAPAK", "HIZLI_BASLANGIC", "ORNEK_VERI", "TESTLER", "AKIS", "DEGISIKLIK_KAYDI"
]

ordered_sheets = [s for s in process_order if s in wb.sheetnames]
remaining_sheets = [s for s in wb.sheetnames if s not in ordered_sheets]
wb._sheets = [wb[s] for s in (ordered_sheets + remaining_sheets)]
print(f"✓ Sheets reordered successfully ({len(wb._sheets)} sheets).")

print("2. Enhancing Tab Color Coding according to Process Stages...")
tab_colors = {
    "KILAVUZ": "059669",
    "PANO": "1E293B",
    "GİRDİLER": "D97706",
    "BANKALAR": "60A5FA", "LIMITLER": "60A5FA", "KREDILER": "60A5FA", "TAKSIT_PLANI": "60A5FA", 
    "ODEME_TAKVIMI": "60A5FA", "MASRAF_KOMISYON": "60A5FA", "TEMINATLAR": "60A5FA", "YENI_TEKLIFLER": "60A5FA",
    "REFINANSMAN": "7C3AED", "YÖNETİM_RAPORU": "7C3AED", "RAPOR": "7C3AED",
    "KARAR_MOTORU": "475569", "MOTOR": "475569", "SENARYO_DUYARLILIK": "475569",
    "KONTROLLER": "64748B", "AYARLAR": "64748B", "LISTELER": "64748B",
    "00_BASLANGIC": "C6A15B", "01_YURUTME_CEKIRDEGI": "065F46", "02_KONTROL_LISTESI": "9B2C2C",
    "03_GORSEL_TASARIM": "475569", "04_MUHASEBE_THP": "2F75B5", "05_KARAR_SATIS": "5B477E"
}

for sname, col in tab_colors.items():
    if sname in wb.sheetnames:
        wb[sname].sheet_properties.tabColor = col

print("✓ Tab colors applied cleanly.")

print("3. Building Comprehensive 5-Stage Process Workflow in KILAVUZ...")
if "KILAVUZ" in wb.sheetnames:
    wb.remove(wb["KILAVUZ"])

ws_k = wb.create_sheet("KILAVUZ", 0)
ws_k.sheet_properties.tabColor = "059669"
ws_k.views.sheetView[0].showGridLines = False

# Margin Col A width: 3.5
ws_k.column_dimensions['A'].width = 3.5
ws_k.column_dimensions['B'].width = 16.0
ws_k.column_dimensions['C'].width = 24.0
ws_k.column_dimensions['D'].width = 46.0
ws_k.column_dimensions['E'].width = 22.0
ws_k.column_dimensions['F'].width = 26.0

ws_k.row_dimensions[1].height = 14.0
ws_k.row_dimensions[2].height = 36.0
ws_k.row_dimensions[3].height = 18.0
ws_k.row_dimensions[4].height = 20.0
ws_k.row_dimensions[5].height = 12.0
ws_k.row_dimensions[6].height = 26.0

# Title
ws_k["B2"] = "BANKA, KREDİ VE TAKSİT TAKİP SİSTEMİ — ÇALIŞMA PROTOKOLÜ"
ws_k["B2"].font = Font(name=FONT_DISPLAY, size=16, bold=True, color=C_NAVY_DARK)

ws_k["B3"] = "C-Level Borç Portföyü, Refinansman Arbitrajı ve Likidite Karar Kokpiti · v15.1 Enterprise Standardı"
ws_k["B3"].font = Font(name=FONT_BODY, size=9.5, italic=True, color=C_TEXT_MUTED)

ws_k["B4"] = "Kurumsal Standart: VUK 185-198, TTK 64, IFRS 9 Finansal Araçlar & Basel III Likidite Kriterleri ile %100 Uyumlu"
ws_k["B4"].font = Font(name=FONT_BODY, size=8.5, bold=True, color=C_EMERALD_TXT)

# Process Table Header (Row 6)
headers = ["AŞAMA NO", "AŞAMA ADI & KAPSAM", "İŞLEM DETAYI VE MANTIK", "KULLANILAN SAYFALAR", "DOĞRUDAN ERİŞİM BUTONU"]
for c_i, h_text in enumerate(headers, start=2):
    c = ws_k.cell(6, c_i, h_text)
    c.font = Font(name=FONT_BODY, size=9, bold=True, color=C_WHITE)
    c.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border_header

# 5 Process Steps
workflow_steps = [
    ("AŞAMA 1", "SİSTEM KURULUMU & PARAMETRELER", 
     "Şirket unvanı, raporlama tarihi, serbest nakit tutarı ve banka risk tolerans eşiklerini AYARLAR sekmesinden tanımlayın. Formüllere asla sabit sayı gömülmez.",
     "AYARLAR, LISTELER", "⚙ AYARLARI YAP (#AYARLAR!A1)", "AYARLAR!A1", C_WHITE),
    ("AŞAMA 2", "STRATEJİK KOKPİT & İZLEME", 
     "Tüm kredi portföyünün anlık büyüklüğünü, 30 günlük ödeme baskısını, likidite erken uyarı radarı ve refinansman aylık tasarruf fırsatını tek ekrandan izleyin.",
     "PANO", "▶ PANO KOKPİTİ (#PANO!A1)", "PANO!A1", C_CARD_BG),
    ("AŞAMA 3", "PORTFÖY VE TAKSİT GİRİŞİ", 
     "Banka limitlerini, çekilen kredileri ve sözleşme amortisman takvimlerini ilgili sayfalara girin. Sarı dolgulu alanlar kullanıcı girdisi, beyaz/mavi alanlar canlı formüldür.",
     "GİRDİLER, KREDILER, TAKSIT_PLANI", "📋 GİRİŞLERİ DÜZENLE (#GİRDİLER!A1)", "GİRDİLER!A1", C_WHITE),
    ("AŞAMA 4", "REFİNANSMAN VE STRES TESTİ", 
     "Faizler düştüğünde krediyi başka bankaya taşımanın erken kapama cezası ve net nakit kazancını REFINANSMAN sayfasında hesaplayın. -%20 nakit daralması stres testini inceleyin.",
     "REFINANSMAN, YÖNETİM_RAPORU", "📊 STRES TESTİ RAPORU (#YÖNETİM_RAPORU!A1)", "YÖNETİM_RAPORU!A1", C_CARD_BG),
    ("AŞAMA 5", "YÖNETİM KURULU VE BANKA ÇIKTISI", 
     "Kredi komitelerine, genel müdüre ve finansal kuruluşlara sunulacak A4 dikey baskıya hazır kurumsal 'Finansal Borçluluk ve Kredi İtibar Raporu' çıktısını alın.",
     "RAPOR", "📄 A4 RAPORU AÇ (#RAPOR!A1)", "RAPOR!A1", C_WHITE)
]

for idx, (st_no, st_name, st_desc, st_sheets, btn_lbl, target, row_bg) in enumerate(workflow_steps, start=7):
    ws_k.row_dimensions[idx].height = 42.0
    
    # Col B: Step No
    c_b = ws_k.cell(idx, 2, st_no)
    c_b.font = Font(name=FONT_BODY, size=9.5, bold=True, color=C_BLUE_TXT)
    c_b.alignment = Alignment(horizontal="center", vertical="center")
    
    # Col C: Step Name
    c_c = ws_k.cell(idx, 3, st_name)
    c_c.font = Font(name=FONT_BODY, size=9, bold=True, color=C_TEXT_MAIN)
    c_c.alignment = Alignment(horizontal="left", vertical="center")
    
    # Col D: Step Desc
    c_d = ws_k.cell(idx, 4, st_desc)
    c_d.font = Font(name=FONT_BODY, size=8.5, color=C_TEXT_SUB)
    c_d.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    # Col E: Sheets
    c_e = ws_k.cell(idx, 5, st_sheets)
    c_e.font = Font(name=FONT_BODY, size=8.5, italic=True, color=C_TEXT_MUTED)
    c_e.alignment = Alignment(horizontal="center", vertical="center")
    
    # Col F: Button
    c_f = ws_k.cell(idx, 6, btn_lbl)
    c_f.font = Font(name=FONT_BODY, size=8.5, bold=True, color=C_WHITE)
    c_f.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
    c_f.alignment = Alignment(horizontal="center", vertical="center")
    c_f.hyperlink = f"#{target}"
    
    for c_i in range(2, 7):
        cell = ws_k.cell(idx, c_i)
        if c_i != 6:
            cell.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
        cell.border = border_thin

# Section 2: Tab Color Code Explanations
ws_k.row_dimensions[13].height = 14.0
ws_k.row_dimensions[14].height = 26.0

ws_k["B14"] = "SİSTEM SEKME RENKLERİ VE MİMARİ ANLAMLARI"
ws_k["B14"].font = Font(name=FONT_DISPLAY, size=11, bold=True, color=C_NAVY_DARK)

color_explanations = [
    ("Yeşil Sekme (KILAVUZ)", "Sistem çalışma prensiplerini ve 5 adımlı proses haritasını içeren ana giriş noktası."),
    ("Koyu Lacivert (PANO)", "C-Level yönetim kokpiti, anlık KPI kartları ve likidite erken uyarı radarı."),
    ("Kehribar Sarı (GİRDİLER)", "Kullanıcının her gün veri girdiği, formülleri besleyen ana operasyon havuzu."),
    ("Açık Mavi (BANKALAR, KREDILER, TAKSIT)", "Banka sözleşmeleri, limitleri ve amortisman takvimlerini saklayan veri katmanı."),
    ("Mor (REFINANSMAN, YÖNETİM RAPORU)", "Stres testleri, arbitraj kazanç hesaplamaları ve yönetim kurulu karar paketleri."),
    ("Koyu Gri (MOTOR, KARAR MOTORU)", "Arka planda çalışan, kullanıcının müdahale etmediği matematiksel hesaplama çekirdeği."),
    ("Kilitli Master Sekmeler (00-05)", "ISO/VUK/IFRS uyumluluğunu, kalite kontrol kapılarını ve yasal sertifikasyonu koruyan katman.")
]

for idx, (c_name, c_meaning) in enumerate(color_explanations, start=15):
    ws_k.row_dimensions[idx].height = 22.0
    c_title = ws_k.cell(idx, 2, c_name)
    c_title.font = Font(name=FONT_BODY, size=8.5, bold=True, color=C_TEXT_MAIN)
    c_title.alignment = Alignment(horizontal="left", vertical="center")
    
    ws_k.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=6)
    c_desc = ws_k.cell(idx, 3, c_meaning)
    c_desc.font = Font(name=FONT_BODY, size=8.5, color=C_TEXT_SUB)
    c_desc.alignment = Alignment(horizontal="left", vertical="center")
    
    for c_i in range(2, 7):
        ws_k.cell(idx, c_i).border = border_thin

print("✓ KILAVUZ 5-stage process workflow and color dictionary built.")

wb.save(wb_path)
print("✓ Workbook saved cleanly.")
