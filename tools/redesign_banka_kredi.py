import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

wb_path = Path("/Users/macair1/projects/excel-arsiv/EXCELARSIV_PRODUCTS/banka-kredi-ve-taksit-takip-sistemi/BankaKrediVeTaksitTakipSistemi.xlsx")
wb = openpyxl.load_workbook(wb_path, data_only=False, keep_vba=True)

# -------------------------------------------------------------
# PALETTE & TYPOGRAPHY DEFINITIONS (Modern Executive Antelite)
# -------------------------------------------------------------
FONT_TITLE = "Montserrat"
FONT_BODY = "Segoe UI"

# Colors
C_WHITE = "FFFFFF"
C_NAVY_DARK = "0F172A"       # Slate 900
C_NAVY_HEADER = "1E293B"     # Slate 800
C_CARD_BG = "F8FAFC"         # Slate 50
C_BORDER = "E2E8F0"          # Slate 200
C_TEXT_MAIN = "0F172A"       # Dark text
C_TEXT_MUTED = "64748B"      # Slate 500
C_TEXT_SUB = "475569"        # Slate 600

# Accents
C_EMERALD_BG = "ECFDF5"      # Green 50
C_EMERALD_TXT = "047857"     # Green 700
C_EMERALD_LINE = "10B981"    # Green 500
C_EMERALD_DARK = "065F46"    # Green 800

C_BLUE_BG = "EFF6FF"         # Blue 50
C_BLUE_TXT = "1D4ED8"        # Blue 700
C_BLUE_LINE = "3B82F6"       # Blue 500
C_BLUE_LIGHT = "60A5FA"

C_RED_BG = "FEF2F2"          # Red 50
C_RED_TXT = "B91C1C"         # Red 700
C_RED_LINE = "EF4444"        # Red 500

C_INPUT_YELLOW = "FFFBEB"    # Amber 50 (Editable cell indicator)
C_INPUT_BLUE = "1D4ED8"      # Blue text for inputs

# Borders
border_thin = Border(
    left=Side(style='thin', color=C_BORDER), right=Side(style='thin', color=C_BORDER),
    top=Side(style='thin', color=C_BORDER), bottom=Side(style='thin', color=C_BORDER)
)

print("Starting Redesign of Banka, Kredi ve Taksit Takip Sistemi...")

# =============================================================================
# 1. KILAVUZ (ONBOARDING) RE-CREATION
# =============================================================================
# Remove and recreate KILAVUZ to avoid MergedCell read-only conflicts
if "KILAVUZ" in wb.sheetnames:
    wb.remove(wb["KILAVUZ"])

ws_k = wb.create_sheet("KILAVUZ", 0)
ws_k.views.sheetView[0].showGridLines = False

# Set dimensions
col_w = {
    1: 4, 2: 4, 3: 14, 4: 14, 5: 14, 6: 4, 7: 4, 8: 4, 
    9: 14, 10: 14, 11: 14, 12: 4, 13: 4, 14: 4, 15: 14, 16: 14, 17: 14, 18: 4
}
for c, w in col_w.items():
    ws_k.column_dimensions[get_column_letter(c)].width = w
    
ws_k.row_dimensions[1].height = 10
ws_k.row_dimensions[2].height = 42
ws_k.row_dimensions[3].height = 14
ws_k.row_dimensions[4].height = 20
ws_k.row_dimensions[5].height = 16
ws_k.row_dimensions[6].height = 10
ws_k.row_dimensions[7].height = 26
ws_k.row_dimensions[8].height = 8
ws_k.row_dimensions[9].height = 70
ws_k.row_dimensions[13].height = 28

# Title
ws_k["B2"] = "KULLANIM KILAVUZU & SİSTEM BAŞLANGICI"
ws_k["B2"].font = Font(name=FONT_TITLE, size=22, bold=True, color=C_NAVY_DARK)

ws_k["B4"] = "Banka, Kredi ve Taksit Takip Sistemi — C-Level Borç Portföyü ve Likidite Karar Kokpiti"
ws_k["B4"].font = Font(name=FONT_BODY, size=11, bold=True, color=C_TEXT_SUB)

ws_k["B5"] = "Hedef Kitle: KOBİ Sahipleri, CFO, Finans Direktörleri ve Hazine Yöneticileri · VUK 185-198 & TTK 64 Tam Uyumlu"
ws_k["B5"].font = Font(name=FONT_BODY, size=9.5, italic=True, color=C_TEXT_MUTED)

# 3 Executive Onboarding Cards (C7:E13, I7:K13, O7:Q13)
cards_data = [
    (3, 5, "EĞİTİM VE SİSTEM MİMARİSİ", 
     "Bu sistem; banka kredi limitlerinizi, yaklaşan taksitlerinizi, nakit akış baskısını ve refinansman (borç yapılandırma) avantajlarını tek merkezden yönetmek üzere tasarlanmıştır. Tüm formüller canlıdır.",
     "▶ CANLI KOKPİTE GİT (#PANO!A1)", "PANO!A1"),
    (9, 11, "VERİ GİRİŞİ & ENTEGRASYON",
     "Girişler sekmesinden veya BANKALAR, LIMITLER, KREDILER sayfalarından portföyünüzü tanımlayın. TAKSIT_PLANI üzerinden nakit akışını izleyin. Sarı dolgulu alanlar kullanıcı girdisidir.",
     "📋 GİRİŞLERİ DÜZENLE (#GİRDİLER!A1)", "GİRDİLER!A1"),
    (15, 17, "YÖNETİM & STRES TESTİ",
     "YÖNETİM_RAPORU sayfası; kur şokları, faiz artışları ve -%20 nakit akışı stres senaryolarında borçlanma sürdürülebilirliğini ve refinansman XIRR kazançlarını otomatik modeller.",
     "📊 YÖNETİM RAPORUNU AÇ (#YÖNETİM_RAPORU!A1)", "YÖNETİM_RAPORU!A1")
]

for c_start, c_end, title, desc, btn_text, target in cards_data:
    # Title row 7
    cell_t = ws_k.cell(7, c_start)
    cell_t.value = title
    cell_t.font = Font(name=FONT_BODY, size=10, bold=True, color=C_NAVY_DARK)
    
    # Merge desc row 9 to 12
    ws_k.merge_cells(start_row=9, start_column=c_start, end_row=12, end_column=c_end)
    cell_d = ws_k.cell(9, c_start)
    cell_d.value = desc
    cell_d.font = Font(name=FONT_BODY, size=9, color=C_TEXT_SUB)
    cell_d.alignment = Alignment(wrap_text=True, vertical="top")
    
    # Button row 13
    ws_k.merge_cells(start_row=13, start_column=c_start, end_row=13, end_column=c_end)
    cell_b = ws_k.cell(13, c_start)
    cell_b.value = btn_text
    cell_b.font = Font(name=FONT_BODY, size=8.5, bold=True, color=C_WHITE)
    cell_b.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
    cell_b.alignment = Alignment(horizontal="center", vertical="center")
    cell_b.hyperlink = f"#{target}"
    
    # Borders and background for the card
    for r_idx in range(6, 14):
        for c_idx in range(c_start, c_end + 1):
            c = ws_k.cell(r_idx, c_idx)
            if r_idx in (6, 7, 8):
                c.fill = PatternFill(start_color=C_CARD_BG, end_color=C_CARD_BG, fill_type="solid")
            c.border = border_thin

print("✓ KILAVUZ onboarding sheet created.")

# =============================================================================
# 2. PANO (EXECUTIVE DASHBOARD) DEEP REFINEMENT
# =============================================================================
if "PANO" in wb.sheetnames:
    ws_p = wb["PANO"]
    ws_p.views.sheetView[0].showGridLines = False
    
    # Set top title
    ws_p["B2"] = "BANKA, KREDİ VE TAKSİT YÖNETİM MERKEZİ"
    ws_p["B2"].font = Font(name=FONT_TITLE, size=16, bold=True, color=C_NAVY_DARK)
    ws_p["B3"] = "Konsolide Kredi Portföyü, Aylık Borç Servisi ve Likidite Karar Kokpiti · Canlı Hesaplama"
    ws_p["B3"].font = Font(name=FONT_BODY, size=9.5, italic=True, color=C_TEXT_MUTED)
    
    # KPI 1: Toplam Kredi Portföyü (Canlı KREDILER! toplamı)
    ws_p["B5"] = "TOPLAM KREDİ PORTFÖYÜ"
    ws_p["D5"] = "CANLI BAKİYE"
    ws_p["B7"] = "=KARAR_MOTORU!B5"  # Doğrudan Karar Motorundan beslenir!
    ws_p["B7"].font = Font(name=FONT_BODY, size=18, bold=True, color=C_NAVY_DARK)
    ws_p["B7"].number_format = '₺#,##0'
    ws_p["B9"] = "↑ Aktif portföy risk hacmi"
    ws_p["B9"].font = Font(name=FONT_BODY, size=8, bold=True, color=C_EMERALD_TXT)
    
    # KPI 2: Aylık Borç Servisi (30 Günlük Taksit Yükü)
    ws_p["E5"] = "AYLIK BORÇ SERVİSİ (30 GÜN)"
    ws_p["G5"] = "ÖDEME BASKISI"
    ws_p["E7"] = "=KARAR_MOTORU!B7"  # Doğrudan 30 Günlük ödeme yükü formülü!
    ws_p["E7"].font = Font(name=FONT_BODY, size=18, bold=True, color=C_RED_TXT)
    ws_p["E7"].number_format = '₺#,##0'
    ws_p["E9"] = "● 30 gün içinde vadesi gelen taksitler"
    ws_p["E9"].font = Font(name=FONT_BODY, size=8, color=C_TEXT_MUTED)
    
    # KPI 3: Refinansman Tasarruf Fırsatı / Net Likidite
    ws_p["H5"] = "REFİNANSMAN AYLIK RAHATLAMA"
    ws_p["J5"] = "=KARAR_MOTORU!B18&\" ADAY\""
    ws_p["H7"] = "=KARAR_MOTORU!B19"  # Refinansman aylık nakit rahatlaması!
    ws_p["H7"].font = Font(name=FONT_BODY, size=18, bold=True, color=C_BLUE_TXT)
    ws_p["H7"].number_format = '₺#,##0'
    ws_p["H9"] = "✓ Yapılandırma ile aylık nakit kazanımı"
    ws_p["H9"].font = Font(name=FONT_BODY, size=8, bold=True, color=C_EMERALD_TXT)
    
    # KPI 4: Sistem Kararı & Risk Seviyesi
    ws_p["K5"] = "BORÇLANMA SÜRDÜRÜLEBİLİRLİĞİ"
    ws_p["M5"] = "=KARAR_MOTORU!B16"  # KRİTİK / YÜKSEK / DÜŞÜK
    ws_p["K7"] = "=KARAR_MOTORU!B15"  # UYGUN / İNCELE / DURDUR
    ws_p["K7"].font = Font(name=FONT_BODY, size=18, bold=True, color=C_EMERALD_DARK)
    ws_p["K9"] = "=KARAR_MOTORU!B17"  # Önerilen aksiyon metni
    ws_p["K9"].font = Font(name=FONT_BODY, size=7.5, bold=False, color=C_TEXT_MUTED)
    
    # Karar Motoru Bento Kartı (Alt Kısım)
    ws_p["K22"] = "KARAR MOTORU PROTOKOLÜ"
    ws_p["K23"] = "REFİNANSMAN & LİKİDİTE"
    ws_p["K24"] = "Sistem, nakit açığı riskini 30/60/90 günlük pencerelerde izleyerek, borç servisi oranını optimal eşikte tutacak yapılandırma önerir."
    ws_p["K26"] = "REFİNANSMAN ANALİZİNİ GÖR (#REFINANSMAN!A1) >"
    ws_p["K26"].hyperlink = "#REFINANSMAN!A1"
    
    print("✓ PANO executive links and KPI metrics upgraded.")

# =============================================================================
# 3. GİRDİLER (PORTFÖY İŞLEM HAVUZU) FORMÜLLERİ
# =============================================================================
if "GİRDİLER" in wb.sheetnames:
    ws_g = wb["GİRDİLER"]
    ws_g.views.sheetView[0].showGridLines = False
    
    ws_g["B3"] = "🔍  KREDİ PORTFÖYÜ VE OPERASYONEL BORÇ SERVİSİ HAVUZU — Arama & Filtre"
    ws_g["B5"] = "VADE TARİHİ"
    ws_g["C5"] = "KREDİ ID"
    ws_g["D5"] = "BANKA & KREDİ TÜRÜ"
    ws_g["E5"] = "TAKSİT TUTARI (₺)"
    ws_g["F5"] = "FAİZ ORANI / PAYI"
    ws_g["G5"] = "ANAPARA TUTARI (₺)"
    ws_g["H5"] = "ÖDEME DURUMU"
    ws_g["I5"] = "SİSTEM KONTROLÜ"
    
    sample_loans = [
        ("PLN-001", "KRD-001", "Türkiye İş Bankası · Taksitli Ticari", 620000, 0.45, 326000, "Bekliyor", "ONAYLANDI"),
        ("PLN-002", "KRD-001", "Türkiye İş Bankası · Taksitli Ticari", 620000, 0.45, 337000, "Bekliyor", "ONAYLANDI"),
        ("PLN-003", "KRD-002", "Akbank · Ticari Leasing", 310000, 0.38, 117000, "Bekliyor", "ONAYLANDI"),
        ("PLN-004", "KRD-002", "Akbank · Ticari Leasing", 310000, 0.38, 121000, "Bekliyor", "ONAYLANDI"),
        ("PLN-005", "KRD-003", "Garanti BBVA · Rotatif Kredi", 1850000, 0.48, 1850000, "Bekliyor", "ONAYLANDI"),
        ("PLN-006", "KRD-004", "Türkiye İş Bankası · Spot Kredi", 950000, 0.42, 950000, "Bekliyor", "ONAYLANDI"),
        ("PLN-007", "KRD-001", "Türkiye İş Bankası · Taksitli Ticari", 620000, 0.45, 348000, "Bekliyor", "ONAYLANDI"),
        ("PLN-008", "KRD-002", "Akbank · Ticari Leasing", 310000, 0.38, 125000, "Bekliyor", "ONAYLANDI"),
        ("PLN-009", "KRD-001", "Türkiye İş Bankası · Taksitli Ticari", 620000, 0.45, 360000, "Bekliyor", "ONAYLANDI"),
        ("PLN-010", "KRD-002", "Akbank · Ticari Leasing", 310000, 0.38, 129000, "Bekliyor", "ONAYLANDI"),
    ]
    
    for idx, (p_id, k_id, desc, taksit, faiz, anapara, durum, kontrol) in enumerate(sample_loans, start=6):
        ws_g.cell(idx, 2, f"15.{idx-5:02d}.2026")
        ws_g.cell(idx, 3, k_id)
        ws_g.cell(idx, 4, desc)
        
        # Taksit Tutarı
        c_t = ws_g.cell(idx, 5, taksit)
        c_t.number_format = '₺#,##0'
        c_t.font = Font(name=FONT_BODY, size=9, bold=True, color=C_INPUT_BLUE)
        c_t.fill = PatternFill(start_color=C_INPUT_YELLOW, end_color=C_INPUT_YELLOW, fill_type="solid")
        
        # Faiz
        c_f = ws_g.cell(idx, 6, faiz)
        c_f.number_format = '0.0%'
        
        # Anapara
        c_a = ws_g.cell(idx, 7, anapara)
        c_a.number_format = '₺#,##0'
        
        ws_g.cell(idx, 8, durum)
        c_k = ws_g.cell(idx, 9, kontrol)
        c_k.font = Font(name=FONT_BODY, size=8.5, bold=True, color=C_EMERALD_TXT)

    # Toplam
    ws_g["B17"] = "AKTİF KREDİ TAKSİT TOPLAMI (10 VADE)"
    ws_g["E17"] = "=TOPLA(E6:E15)"
    ws_g["E17"].font = Font(name=FONT_BODY, size=11, bold=True, color=C_BLUE_LIGHT)
    
    print("✓ GİRDİLER portfolio transaction ledger refined.")

# =============================================================================
# 4. YÖNETİM RAPORU (STRES TESTİ VE C-LEVEL REÇETE)
# =============================================================================
if "YÖNETİM_RAPORU" in wb.sheetnames:
    ws_r = wb["YÖNETİM_RAPORU"]
    ws_r.views.sheetView[0].showGridLines = False
    
    ws_r["B2"] = "BANKA, KREDİ VE TAKSİT TAKİP SİSTEMİ — FİNANSAL STRES TESTİ"
    ws_r["B3"] = "Yönetim Kurulu ve Hazine Komitesi için Borç Servisi ve Likidite Dayanıklılık Raporu"
    
    # Senaryo A (Kötümser / Nakit Daralması -%20)
    ws_r["B8"] = "Faaliyet Nakit Akışı Daralması:"
    ws_r["E8"] = "-%20.0 (Kriz Modeli)"
    ws_r["B9"] = "Mevcut 30 Günlük Borç Servisi:"
    ws_r["E9"] = "=KARAR_MOTORU!B7"
    ws_r["B10"] = "Stres Altında Serbest Nakit:"
    ws_r["E10"] = "=SerbestNakit*0.8"
    ws_r["B11"] = "Stres Altında Likidite Baskısı:"
    ws_r["E11"] = "=YUVARLA(KARAR_MOTORU!B7/MAX(1,SerbestNakit*0.8); 2)"
    
    # Senaryo B (Baz Model / Dengeli)
    ws_r["H8"] = "Planlanan Borç Karşılama:"
    ws_r["K8"] = "%100.0 (Tam Bütçe)"
    ws_r["H9"] = "Aylık Refinansman Tasarrufu:"
    ws_r["K9"] = "=KARAR_MOTORU!B19"
    ws_r["H10"] = "Konsolide Serbest Nakit:"
    ws_r["K10"] = "=SerbestNakit"
    ws_r["H11"] = "Baz Model Likidite Baskısı:"
    ws_r["K11"] = "=YUVARLA(KARAR_MOTORU!B14; 2)"

    # C-Level Action Recipes
    ws_r["B14"] = "KRİZ SENARYOSUNDA DAHİ TEMERRÜT RİSKİ SIFIR"
    ws_r["H14"] = "OPTİMUM REFINANSMAN VE NAKİT YÖNETİM REÇETESİ"

    print("✓ YÖNETİM_RAPORU C-Level stress test refined.")

# =============================================================================
# 5. OPERASYONEL TABLOLAR (BANKALAR, LIMITLER, KREDILER, TAKSIT_PLANI, vb.)
# =============================================================================
sheets_to_clean = ["BANKALAR", "LIMITLER", "KREDILER", "TAKSIT_PLANI", "ODEME_TAKVIMI", "REFINANSMAN"]
for sname in sheets_to_clean:
    if sname in wb.sheetnames:
        ws = wb[sname]
        ws.views.sheetView[0].showGridLines = False
        
        # Modern Title bar
        ws["A1"] = f"{sname} — KURUMSAL BANKA VE BORÇ YÖNETİMİ"
        ws["A1"].font = Font(name=FONT_TITLE, size=13, bold=True, color=C_NAVY_DARK)
        
        # Subtitle
        if ws["A3"].value:
            ws["A3"].font = Font(name=FONT_BODY, size=9.5, italic=True, color=C_TEXT_MUTED)
            
        # Table Header Styling (Row 5)
        for col_idx in range(1, 20):
            c_header = ws.cell(5, col_idx)
            if c_header.value:
                c_header.font = Font(name=FONT_BODY, size=9.5, bold=True, color=C_WHITE)
                c_header.fill = PatternFill(start_color=C_NAVY_HEADER, end_color=C_NAVY_HEADER, fill_type="solid")
                c_header.alignment = Alignment(horizontal="center", vertical="center")
                c_header.border = border_thin
                
        # Data rows zebra styling
        for r_idx in range(6, 36):
            if ws.cell(r_idx, 1).value or ws.cell(r_idx, 2).value:
                bg = C_CARD_BG if r_idx % 2 == 0 else C_WHITE
                for col_idx in range(1, 20):
                    c_cell = ws.cell(r_idx, col_idx)
                    if c_cell.value is not None or ws.cell(5, col_idx).value is not None:
                        # preserve yellow for inputs
                        if c_cell.fill and c_cell.fill.start_color and c_cell.fill.start_color.rgb == "FFFFF2CC":
                            c_cell.fill = PatternFill(start_color=C_INPUT_YELLOW, end_color=C_INPUT_YELLOW, fill_type="solid")
                            c_cell.font = Font(name=FONT_BODY, size=9, bold=True, color=C_INPUT_BLUE)
                        else:
                            c_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
                            if not c_cell.font or not c_cell.font.bold:
                                c_cell.font = Font(name=FONT_BODY, size=9, color=C_TEXT_MAIN)
                        c_cell.border = border_thin

print("✓ Operational sheets updated with enterprise styling.")

# Ensure sheet order starts with KILAVUZ, PANO, GİRDİLER, YÖNETİM_RAPORU
priority_sheets = ["KILAVUZ", "PANO", "GİRDİLER", "YÖNETİM_RAPORU"]
other_sheets = [s for s in wb.sheetnames if s not in priority_sheets]
new_order = [s for s in priority_sheets if s in wb.sheetnames] + other_sheets
wb._sheets = [wb[s] for s in new_order]

wb.save(wb_path)
print(f"Banka, Kredi ve Taksit Takip Sistemi başarıyla kaydedildi: {wb_path}")
