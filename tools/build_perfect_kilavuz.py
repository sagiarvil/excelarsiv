import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

wb_path = Path("/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx")
wb = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=False)

if "KILAVUZ" in wb.sheetnames:
    wb.remove(wb["KILAVUZ"])

ws = wb.create_sheet("KILAVUZ", 0)
ws.sheet_properties.tabColor = "059669"
ws.views.sheetView[0].showGridLines = False

# Colors
C_WHITE = "FFFFFF"
C_CARD_FRAME = "F2F2F2"  # Exact reference card color (Theme 0, tint -0.05)
C_CARD_HEADER = "E6E6E6" # Card header border
C_TEXT_MAIN = "000000"
C_TEXT_CARD_TITLE = "000000"
C_TEXT_CARD_BODY = "333333"
C_BORDER_GRAY = "CCCCCC"
C_NAVY_BTN = "1B365D"    # Executive Navy Button
C_LINE_ACCENT = "059669"

# Font names
FONT_TITLE = "Montserrat"
FONT_CARD_TITLE = "Neue Haas Grotesk Text Proalibr"
FONT_BODY = "Century Gothic"

# 1. Exact Column Dimensions from Reference
col_widths = {
    'A': 3.6,
    'B': 3.7, 'C': 10.4, 'D': 13.0, 'E': 13.0, 'F': 3.7,  # Card 1 [B..F] Total width ~ 43.8
    'G': 8.8,                                              # Gap 1
    'H': 3.7, 'I': 10.4, 'J': 13.0, 'K': 13.0, 'L': 3.7,  # Card 2 [H..L] Total width ~ 43.8
    'M': 7.5,                                              # Gap 2
    'N': 3.7, 'O': 10.4, 'P': 13.0, 'Q': 13.0, 'R': 3.7,  # Card 3 [N..R] Total width ~ 43.8
    'S': 8.8
}
for col_l, w in col_widths.items():
    ws.column_dimensions[col_l].width = w

# 2. Row Heights from Reference
ws.row_dimensions[1].height = 7.8
ws.row_dimensions[2].height = 36.0  # Title row
ws.row_dimensions[3].height = 7.8
ws.row_dimensions[4].height = 17.4
ws.row_dimensions[5].height = 17.4
ws.row_dimensions[6].height = 10.0  # Top of cards
ws.row_dimensions[7].height = 24.0  # Card titles
ws.row_dimensions[8].height = 8.0   # Title underline
ws.row_dimensions[9].height = 20.0  # Card body
ws.row_dimensions[10].height = 20.0
ws.row_dimensions[11].height = 20.0
ws.row_dimensions[12].height = 20.0
ws.row_dimensions[13].height = 10.0
ws.row_dimensions[14].height = 28.0 # Action Button
ws.row_dimensions[15].height = 8.0
ws.row_dimensions[16].height = 8.0

# 3. Main Title in B2:M2
ws.merge_cells("B2:M2")
ws["B2"] = "KILAVUZ"
ws["B2"].font = Font(name=FONT_TITLE, size=24, bold=True, color="1F2937")
ws["B2"].alignment = Alignment(horizontal="left", vertical="center")

# Subtitle in B4
ws["B4"] = "Banka, Kredi ve Taksit Takip Sistemi — C-Level Borç Portföyü ve Likidite Karar Kokpiti"
ws["B4"].font = Font(name="Segoe UI", size=11, bold=True, color="4B5563")

# Subtitle in B5
ws["B5"] = "Hedef Kitle: KOBİ Sahipleri, CFO, Finans Direktörleri ve Hazine Yöneticileri · VUK 185-198 & TTK 64 Tam Uyumlu"
ws["B5"].font = Font(name="Segoe UI", size=9.5, italic=True, color="6B7280")

# 4. Fill Reference Card Blocks [B6:F16], [H6:L16], [N6:R16]
card_specs = [
    (2, 6, "C7", "C9:E12", "C14:E14", "EĞİTİM VİDEOSU", 
     "Bu şablonun nasıl kullanıldığını videolu ve detaylı bir şekilde öğrenmek için aşağıdaki düğmeye tıklayabilirsiniz.",
     "▶ KULLANIM VİDEOSUNU İZLE", "PANO!A1"),
    (8, 12, "I7", "I9:K12", "I14:K14", "KULLANICI KILAVUZU", 
     "Excel Şablonlarımızı nasıl kullanmanız gerektiğini öğrenmek için buraya tıklayarak Kullanıcı Kılavuzunu indirebilirsiniz.",
     "📥 KULLANICI KILAVUZUNU İNDİR", "GİRDİLER!A1"),
    (14, 18, "O7", "O9:Q12", "O14:Q14", "EXCEL ARŞİV HİZMETLERİ", 
     "EXCELARŞİV olarak finansal modelleme, kurumsal danışmanlık ve karar destek sistemleri sunmaktayız. Daha fazlası için tıklayınız.",
     "🌐 WEB SİTEMİZİ ZİYARET EDİN", "YÖNETİM_RAPORU!A1")
]

fill_card = PatternFill(start_color=C_CARD_FRAME, end_color=C_CARD_FRAME, fill_type="solid")
border_bottom_title = Border(bottom=Side(style='thin', color=C_BORDER_GRAY))

for c_start, c_end, title_cell, body_range, btn_range, title_text, body_text, btn_text, target in card_specs:
    # Fill the background block exactly like the reference
    # Outer frame: Rows 6..16, Cols c_start..c_end
    for r in range(6, 17):
        for c in range(c_start, c_end + 1):
            cell = ws.cell(r, c)
            # Center of the card (rows 9..12, inner 3 cols) is white background!
            if r in range(9, 13) and c in (c_start + 1, c_start + 2, c_start + 3):
                cell.fill = PatternFill(fill_type=None)
            else:
                cell.fill = fill_card

    # Card Title
    t_col = c_start + 1
    ws.cell(7, t_col).value = title_text
    ws.cell(7, t_col).font = Font(name=FONT_CARD_TITLE, size=11, bold=True, color=C_TEXT_CARD_TITLE)
    ws.cell(7, t_col).alignment = Alignment(horizontal="left", vertical="center")
    
    # Title underline across inner 3 columns
    for c_i in range(c_start + 1, c_end):
        ws.cell(7, c_i).border = border_bottom_title

    # Card Body (Merged range)
    ws.merge_cells(body_range)
    b_cell = ws[body_range.split(':')[0]]
    b_cell.value = body_text
    b_cell.font = Font(name=FONT_BODY, size=10, bold=False, color=C_TEXT_CARD_BODY)
    b_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Action Button (Merged range)
    ws.merge_cells(btn_range)
    btn_cell = ws[btn_range.split(':')[0]]
    btn_cell.value = btn_text
    btn_cell.font = Font(name="Segoe UI", size=9, bold=True, color=C_WHITE)
    btn_cell.fill = PatternFill(start_color=C_NAVY_BTN, end_color=C_NAVY_BTN, fill_type="solid")
    btn_cell.alignment = Alignment(horizontal="center", vertical="center")
    btn_cell.hyperlink = f"#{target}"

    # Round/bordered styling for button
    b_style = Border(
        left=Side(style='thin', color="0F172A"), right=Side(style='thin', color="0F172A"),
        top=Side(style='thin', color="0F172A"), bottom=Side(style='thin', color="0F172A")
    )
    for c_i in range(c_start + 1, c_end):
        ws.cell(14, c_i).border = b_style

# 5. Section 2: 5-Stage Process Table below the cards (Rows 19+)
ws.row_dimensions[18].height = 14.0
ws.row_dimensions[19].height = 28.0
ws["B19"] = "5 AŞAMALI KURUMSAL ÇALIŞMA PROTOKOLÜ (PROSES REHBERİ)"
ws["B19"].font = Font(name=FONT_TITLE, size=12, bold=True, color="1F2937")

headers = ["AŞAMA NO", "AŞAMA ADI & KAPSAM", "İŞLEM DETAYI VE MANTIK", "KULLANILAN SAYFALAR", "DOĞRUDAN ERİŞİM BUTONU"]
ws.row_dimensions[21].height = 26.0

for c_i, h_text in enumerate(headers, start=2):
    c = ws.cell(21, c_i, h_text)
    c.font = Font(name="Segoe UI", size=9, bold=True, color=C_WHITE)
    c.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")

# Workflow Steps
workflow_steps = [
    ("AŞAMA 1", "SİSTEM KURULUMU & PARAMETRELER", 
     "Şirket unvanı, raporlama tarihi, serbest nakit ve risk toleranslarını AYARLAR sekmesinden tanımlayın.",
     "AYARLAR, LISTELER", "⚙ AYARLAR (#AYARLAR!A1)", "AYARLAR!A1"),
    ("AŞAMA 2", "STRATEJİK KOKPİT & İZLEME", 
     "Kredi portföy büyüklüğü, 30 günlük ödeme baskısı ve likidite erken uyarı radarını tek ekrandan izleyin.",
     "PANO", "▶ PANO KOKPİTİ (#PANO!A1)", "PANO!A1"),
    ("AŞAMA 3", "PORTFÖY VE TAKSİT GİRİŞİ", 
     "Banka limitlerini, çekilen kredileri ve sözleşme amortisman takvimlerini ilgili sayfalara girin.",
     "GİRDİLER, KREDILER, TAKSIT_PLANI", "📋 GİRİŞLERİ DÜZENLE (#GİRDİLER!A1)", "GİRDİLER!A1"),
    ("AŞAMA 4", "REFİNANSMAN VE STRES TESTİ", 
     "Krediyi başka bankaya taşıma durumunda net faiz kârını hesaplayın. -%20 nakit daralması stres testini inceleyin.",
     "REFINANSMAN, YÖNETİM_RAPORU", "📊 STRES TESTİ (#YÖNETİM_RAPORU!A1)", "YÖNETİM_RAPORU!A1"),
    ("AŞAMA 5", "YÖNETİM KURULU VE BANKA ÇIKTISI", 
     "Kredi komiteleri ve yönetim kuruluna sunulacak A4 dikey baskıya hazır kurumsal rapor çıktısını alın.",
     "RAPOR", "📄 A4 RAPORU (#RAPOR!A1)", "RAPOR!A1")
]

border_table = Border(
    left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"),
    top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1")
)

# Merge across to fit the wide reference columns
for idx, (st_no, st_name, st_desc, st_sheets, btn_lbl, target) in enumerate(workflow_steps, start=22):
    ws.row_dimensions[idx].height = 36.0
    row_bg = "FFFFFF" if idx % 2 == 0 else "F8FAFC"
    
    # Col B: Step No
    ws.cell(idx, 2, st_no).font = Font(name="Segoe UI", size=9, bold=True, color="1D4ED8")
    ws.cell(idx, 2).alignment = Alignment(horizontal="center", vertical="center")
    
    # Col C: Step Name
    ws.cell(idx, 3, st_name).font = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
    ws.cell(idx, 3).alignment = Alignment(horizontal="left", vertical="center")
    
    # Merge D..M for Description
    ws.merge_cells(start_row=idx, start_column=4, end_row=idx, end_column=13)
    ws.cell(idx, 4, st_desc).font = Font(name="Segoe UI", size=8.5, color="475569")
    ws.cell(idx, 4).alignment = Alignment(horizontal="left", vertical="center")
    
    # Merge N..O for Sheets
    ws.merge_cells(start_row=idx, start_column=14, end_row=idx, end_column=15)
    ws.cell(idx, 14, st_sheets).font = Font(name="Segoe UI", size=8.5, italic=True, color="64748B")
    ws.cell(idx, 14).alignment = Alignment(horizontal="center", vertical="center")
    
    # Merge P..R for Button
    ws.merge_cells(start_row=idx, start_column=16, end_row=idx, end_column=18)
    b_cell = ws.cell(idx, 16, btn_lbl)
    b_cell.font = Font(name="Segoe UI", size=8.5, bold=True, color=C_WHITE)
    b_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    b_cell.alignment = Alignment(horizontal="center", vertical="center")
    b_cell.hyperlink = f"#{target}"

    for c_i in range(2, 19):
        c_item = ws.cell(idx, c_i)
        if c_i < 16:
            c_item.fill = PatternFill(start_color=row_bg, end_color=row_bg, fill_type="solid")
        c_item.border = border_table

# Ensure tab order: KILAVUZ first
if wb.sheetnames[0] != "KILAVUZ":
    wb._sheets.remove(ws)
    wb._sheets.insert(0, ws)

wb.save(wb_path)
print("✓ Perfect 1:1 Reference KILAVUZ Built and Saved!")
