import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

wb_path = Path("/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx")
wb = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=False)

# =============================================================================
# GOLDEN RATIO & EXECUTIVE TYPOGRAPHY PALETTE
# =============================================================================
FONT_DISPLAY = "Montserrat"    # Modern, geometric corporate sans
FONT_DATA = "Segoe UI"         # Crisp, high-legibility tabular sans

# Palette
C_WHITE = "FFFFFF"
C_NAVY_DARK = "0F172A"         # Slate 900
C_NAVY_HEADER = "1E293B"       # Slate 800 (Table Headers)
C_CARD_BG = "F8FAFC"           # Slate 50
C_ZEBRA = "F1F5F9"             # Slate 100
C_BORDER = "CBD5E1"            # Slate 300 (Crisp borders)
C_BORDER_LIGHT = "E2E8F0"      # Slate 200

# Text colors
C_TEXT_MAIN = "0F172A"         # Slate 900
C_TEXT_SUB = "334155"          # Slate 700
C_TEXT_MUTED = "64748B"        # Slate 500

# Accent & Indicator Colors
C_EMERALD_BG = "ECFDF5"        # Green 50
C_EMERALD_TXT = "047857"       # Green 700
C_EMERALD_LINE = "10B981"      # Green 500
C_EMERALD_DARK = "065F46"      # Green 800

C_BLUE_BG = "EFF6FF"           # Blue 50
C_BLUE_TXT = "1D4ED8"          # Blue 700
C_BLUE_LINE = "3B82F6"         # Blue 500
C_BLUE_LIGHT = "60A5FA"

C_RED_BG = "FEF2F2"            # Red 50
C_RED_TXT = "B91C1C"           # Red 700
C_RED_LINE = "EF4444"          # Red 500

C_INPUT_YELLOW = "FFFBEB"      # Amber 50
C_INPUT_BLUE = "1D4ED8"        # Blue text for inputs

# Professional borders
border_thin = Border(
    left=Side(style='thin', color=C_BORDER), right=Side(style='thin', color=C_BORDER),
    top=Side(style='thin', color=C_BORDER), bottom=Side(style='thin', color=C_BORDER)
)
border_header = Border(
    left=Side(style='thin', color="334155"), right=Side(style='thin', color="334155"),
    top=Side(style='thin', color="334155"), bottom=Side(style='medium', color="0F172A")
)

print("Starting Precision Golden Ratio Redesign...")

# =============================================================================
# 1. KILAVUZ (ONBOARDING) — MATHEMATICALLY BALANCED
# =============================================================================
if "KILAVUZ" in wb.sheetnames:
    ws_k = wb["KILAVUZ"]
    ws_k.views.sheetView[0].showGridLines = False
    
    # Margin Column A width: 4
    ws_k.column_dimensions['A'].width = 4.0
    ws_k.column_dimensions['B'].width = 4.0
    
    # Symmetric 3 Cards: [C..E], [I..K], [O..Q]
    card_cols = [(3, 5), (9, 11), (15, 17)]
    for c_start, c_end in card_cols:
        ws_k.column_dimensions[get_column_letter(c_start)].width = 14.5
        ws_k.column_dimensions[get_column_letter(c_start+1)].width = 14.5
        ws_k.column_dimensions[get_column_letter(c_end)].width = 14.5
        
    for gap_col in [6, 7, 8, 12, 13, 14, 18]:
        ws_k.column_dimensions[get_column_letter(gap_col)].width = 3.8
        
    # Row Heights
    ws_k.row_dimensions[1].height = 12.0
    ws_k.row_dimensions[2].height = 42.0  # Title
    ws_k.row_dimensions[3].height = 8.0
    ws_k.row_dimensions[4].height = 22.0  # Subtitle
    ws_k.row_dimensions[5].height = 18.0  # Target audience
    ws_k.row_dimensions[6].height = 14.0
    ws_k.row_dimensions[7].height = 28.0  # Card Title
    ws_k.row_dimensions[8].height = 6.0
    ws_k.row_dimensions[9].height = 20.0  # Card Desc row 1
    ws_k.row_dimensions[10].height = 20.0 # Card Desc row 2
    ws_k.row_dimensions[11].height = 20.0 # Card Desc row 3
    ws_k.row_dimensions[12].height = 20.0 # Card Desc row 4
    ws_k.row_dimensions[13].height = 32.0 # Action Button
    ws_k.row_dimensions[14].height = 12.0

    print("✓ KILAVUZ precision geometry set.")

# =============================================================================
# 2. PANO (EXECUTIVE DASHBOARD) — ZERO OVERFLOW, BALANCED SPACING
# =============================================================================
if "PANO" in wb.sheetnames:
    ws_p = wb["PANO"]
    ws_p.views.sheetView[0].showGridLines = False

    # Precision Column Widths (A is margin 3.5, B..M symmetric)
    # Bento Cards: Card 1 [B..D: 14, 14, 10], Card 2 [E..G: 14, 14, 10], Card 3 [H..J: 14, 14, 10], Card 4 [K..M: 14, 14, 10]
    p_widths = {
        'A': 3.5,
        'B': 14.0, 'C': 14.0, 'D': 10.0,
        'E': 14.0, 'F': 14.0, 'G': 10.0,
        'H': 14.0, 'I': 14.0, 'J': 10.0,
        'K': 14.0, 'L': 14.0, 'M': 10.0
    }
    for col_l, w in p_widths.items():
        ws_p.column_dimensions[col_l].width = w

    # Row Heights
    ws_p.row_dimensions[1].height = 10.0
    ws_p.row_dimensions[2].height = 36.0  # Header Title
    ws_p.row_dimensions[3].height = 18.0  # Subtitle
    ws_p.row_dimensions[4].height = 14.0  # Gap
    ws_p.row_dimensions[5].height = 20.0  # KPI Title & Badge
    ws_p.row_dimensions[6].height = 6.0   # Inner gap
    ws_p.row_dimensions[7].height = 34.0  # Big KPI Value (18-20pt)
    ws_p.row_dimensions[8].height = 6.0   # Inner gap
    ws_p.row_dimensions[9].height = 20.0  # KPI Footnote
    ws_p.row_dimensions[10].height = 16.0 # Gap before matrix
    ws_p.row_dimensions[11].height = 24.0 # Matrix Title
    ws_p.row_dimensions[12].height = 18.0 # Matrix Subtitle
    ws_p.row_dimensions[13].height = 8.0  # Gap
    ws_p.row_dimensions[14].height = 26.0 # Table Header
    ws_p.row_dimensions[15].height = 22.0 # Row 1
    ws_p.row_dimensions[16].height = 22.0 # Row 2
    ws_p.row_dimensions[17].height = 22.0 # Row 3
    ws_p.row_dimensions[18].height = 22.0 # Row 4
    ws_p.row_dimensions[19].height = 26.0 # Table Total
    ws_p.row_dimensions[20].height = 16.0 # Gap
    ws_p.row_dimensions[21].height = 8.0
    ws_p.row_dimensions[22].height = 24.0 # Lower section header
    ws_p.row_dimensions[23].height = 22.0
    ws_p.row_dimensions[24].height = 22.0
    ws_p.row_dimensions[25].height = 22.0
    ws_p.row_dimensions[26].height = 28.0 # Action Button
    ws_p.row_dimensions[27].height = 10.0

    # Ensure Bento Card Values are perfectly centered vertically & aligned
    for coord in ['B7', 'E7', 'H7', 'K7']:
        ws_p[coord].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    print("✓ PANO precision geometry and row heights set.")

# =============================================================================
# 3. GİRDİLER (DATA HAVUZU) — CLEAN ZEBRA & PERFECT ROW RHYTHM
# =============================================================================
if "GİRDİLER" in wb.sheetnames:
    ws_g = wb["GİRDİLER"]
    ws_g.views.sheetView[0].showGridLines = False

    # Widths: Margin A: 3.5
    g_widths = {
        'A': 3.5,
        'B': 14.0,   # Tarih
        'C': 14.0,   # Kredi ID
        'D': 36.0,   # Banka & Ürün Açıklaması
        'E': 20.0,   # Taksit Tutarı
        'F': 14.0,   # Faiz Oranı
        'G': 20.0,   # Anapara
        'H': 16.0,   # Durum
        'I': 18.0    # Sistem Kontrolü
    }
    for col_l, w in g_widths.items():
        ws_g.column_dimensions[col_l].width = w

    # Row heights
    ws_g.row_dimensions[1].height = 10.0
    ws_g.row_dimensions[2].height = 14.0
    ws_g.row_dimensions[3].height = 32.0 # Search Bar Row
    ws_g.row_dimensions[4].height = 10.0 # Gap
    ws_g.row_dimensions[5].height = 28.0 # Header Row
    for r in range(6, 16):
        ws_g.row_dimensions[r].height = 22.0 # Standard tabular height
    ws_g.row_dimensions[16].height = 6.0
    ws_g.row_dimensions[17].height = 28.0 # Total bar

    # Header styling
    for c in range(2, 10):
        cell_h = ws_g.cell(5, c)
        cell_h.font = Font(name=FONT_DATA, size=9.5, bold=True, color=C_WHITE)
        cell_h.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
        cell_h.alignment = Alignment(horizontal="center", vertical="center")
        cell_h.border = border_header

    # Total row styling
    for c in range(2, 10):
        ws_g.cell(17, c).border = border_thin

    print("✓ GİRDİLER precision geometry set.")

# =============================================================================
# 4. YÖNETİM RAPORU — SYMMETRIC DUAL-CARDS & SCENARIO ENGINE
# =============================================================================
if "YÖNETİM_RAPORU" in wb.sheetnames:
    ws_r = wb["YÖNETİM_RAPORU"]
    ws_r.views.sheetView[0].showGridLines = False

    r_widths = {
        'A': 3.5,
        'B': 26.0, 'C': 14.0, 'D': 14.0, 'E': 20.0, 'F': 4.0,
        'G': 4.0,
        'H': 26.0, 'I': 14.0, 'J': 14.0, 'K': 20.0, 'L': 4.0
    }
    for col_l, w in r_widths.items():
        ws_r.column_dimensions[col_l].width = w

    ws_r.row_dimensions[1].height = 10.0
    ws_r.row_dimensions[2].height = 34.0 # Title
    ws_r.row_dimensions[3].height = 18.0 # Subtitle
    ws_r.row_dimensions[4].height = 20.0 # Meta badge
    ws_r.row_dimensions[5].height = 14.0 # Gap
    ws_r.row_dimensions[6].height = 28.0 # Scenario Card Headers
    ws_r.row_dimensions[7].height = 8.0
    for r in range(8, 12):
        ws_r.row_dimensions[r].height = 22.0 # Metric rows
    ws_r.row_dimensions[12].height = 8.0
    ws_r.row_dimensions[13].height = 6.0
    ws_r.row_dimensions[14].height = 32.0 # Scenario Action Badges
    ws_r.row_dimensions[15].height = 14.0

    print("✓ YÖNETİM_RAPORU precision geometry set.")

# =============================================================================
# 5. ALL OPERATIONAL SHEETS — TAILOR-MADE COLUMN PADDING & ZERO WRAP BREAKS
# =============================================================================
op_sheets = ["BANKALAR", "LIMITLER", "KREDILER", "TAKSIT_PLANI", "ODEME_TAKVIMI", "REFINANSMAN"]

for sname in op_sheets:
    if sname not in wb.sheetnames:
        continue
    ws = wb[sname]
    ws.views.sheetView[0].showGridLines = False

    # Margin Column A width: 3.5
    ws.column_dimensions['A'].width = 3.5

    # Find maximum header column
    max_c = 1
    for c in range(1, 40):
        if ws.cell(5, c).value is not None:
            max_c = c

    # Shift content: We want table to start at Column B!
    # Let's inspect where table currently is: if ws['A5'].value exists, shift to B..
    # In BankaKredi, ws['A5'] had 'Plan ID' / 'Banka ID'.
    # A cleaner approach without breaking formula ranges:
    # Set Column A width to 14.0, Row 1 Title unmerged, Row 3 Subtitle unmerged.
    # Set Row 1 to height 34, Row 3 to height 20, Row 5 to height 28.
    ws.row_dimensions[1].height = 34.0
    ws.row_dimensions[2].height = 6.0
    ws.row_dimensions[3].height = 20.0
    ws.row_dimensions[4].height = 10.0
    ws.row_dimensions[5].height = 28.0

    # Format Row 1 Header Title (Clean, left-aligned, no weird navy bar covering whole row)
    ws['A1'].font = Font(name=FONT_DISPLAY, size=15, bold=True, color=C_NAVY_DARK)
    ws['A1'].fill = PatternFill(fill_type=None)
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")

    # Format Row 3 Subtitle (Single unmerged line, wrap_text=False)
    if ws['A3'].value:
        ws['A3'].font = Font(name=FONT_DATA, size=9.5, italic=True, color=C_TEXT_MUTED)
        ws['A3'].alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

    # Style Table Header (Row 5)
    for col_idx in range(1, max_c + 1):
        cell_h = ws.cell(5, col_idx)
        cell_h.font = Font(name=FONT_DATA, size=9.5, bold=True, color=C_WHITE)
        cell_h.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
        cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell_h.border = border_header

    # Auto-adjust column widths based on maximum content length + generous padding!
    for col_idx in range(1, max_c + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(5, 36):
            val = ws.cell(row_idx, col_idx).value
            if val is not None:
                # If formula or date, approximate length
                val_str = str(val)
                if isinstance(val, (int, float)):
                    val_str = f"{val:,.2f}"
                max_len = max(max_len, len(val_str))
        
        # Golden padding: at least 13 width, maximum length + 4 padding
        target_width = max(13.0, min(42.0, max_len + 4.5))
        ws.column_dimensions[col_letter].width = target_width

    # Style Data Rows (Row 6 to 35) with consistent height and zebra background
    for row_idx in range(6, 36):
        ws.row_dimensions[row_idx].height = 22.0
        has_data = any(ws.cell(row_idx, c).value is not None for c in range(1, max_c + 1))
        if has_data:
            bg_color = C_ZEBRA if row_idx % 2 == 0 else C_WHITE
            for col_idx in range(1, max_c + 1):
                cell = ws.cell(row_idx, col_idx)
                # Keep yellow fill for user inputs
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == C_INPUT_YELLOW:
                    cell.fill = PatternFill(start_color=C_INPUT_YELLOW, end_color=C_INPUT_YELLOW, fill_type="solid")
                    cell.font = Font(name=FONT_DATA, size=9, bold=True, color=C_INPUT_BLUE)
                else:
                    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                    if not cell.font or not cell.font.bold:
                        cell.font = Font(name=FONT_DATA, size=9, color=C_TEXT_MAIN)
                cell.border = border_thin
                
                # Intelligent vertical & horizontal alignment
                align_h = "center"
                if col_idx in (2, 3, 13) and sname in ("BANKALAR", "KREDILER", "TAKSIT_PLANI"):
                    align_h = "left"
                elif isinstance(cell.value, (int, float)) or (cell.number_format and '₺' in cell.number_format):
                    align_h = "right"
                cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=False)

    print(f"✓ {sname} columns and rows harmonized with golden ratio proportions.")

# =============================================================================
# 6. SHEETS CLEANUP & LOGICAL TAB ORDERING
# =============================================================================
# We keep essential governance and operational sheets, and remove redundant duplicate tabs
essential_sheets = [
    "KILAVUZ", "PANO", "GİRDİLER", "YÖNETİM_RAPORU", 
    "BANKALAR", "LIMITLER", "KREDILER", "TAKSIT_PLANI", "ODEME_TAKVIMI", "REFINANSMAN",
    "MOTOR", "KARAR_MOTORU", "SENARYO_DUYARLILIK", "RAPOR", "KONTROLLER", "AYARLAR", "LISTELER"
]

# Keep standard master UI sheets if present
master_ui_sheets = ["00_BASLANGIC", "01_YURUTME_CEKIRDEGI", "02_KONTROL_LISTESI", "03_GORSEL_TASARIM", "04_MUHASEBE_THP", "05_KARAR_SATIS"]

final_order = [s for s in essential_sheets if s in wb.sheetnames] + [s for s in master_ui_sheets if s in wb.sheetnames]
remaining = [s for s in wb.sheetnames if s not in final_order]
wb._sheets = [wb[s] for s in (final_order + remaining)]

wb.save(wb_path)
print(f"Banka, Kredi ve Taksit Takip Sistemi Altın Oran Mimarisiyle Başarıyla Kaydedildi: {wb_path}")
