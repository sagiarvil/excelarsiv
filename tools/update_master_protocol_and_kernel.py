import re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =============================================================================
# 1. UPDATE MANDATE IN EXCELARSIV_DECISION_OS_V15_1_RUNTIME_KERNEL.py
# =============================================================================
kernel_path = Path("/Users/macair1/projects/excel-arsiv/calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_RUNTIME_KERNEL.py")
content = kernel_path.read_text(encoding="utf-8")

ADDITIONAL_MANDATE_RULES = r"""
---

# 249 — GOLDEN RATIO UI, SİMETRİ VE 5 AŞAMALI CFO SÜREÇ PROTOKOLÜ (v15.2 ANAYASASI)

Bu bölüm, her ExcelArşiv şablonunun en az $1,000 ticari değer yaratması ve insan-bilgisayar etkileşimi (HCI) standartlarında kusursuz olması için bağlayıcı kurallardır:

## 249.1 Sıfır Metin Kırılması Kuralı (Zero Text-Wrap & Single-Line Header)
- Tablo üst başlıkları ve alt açıklamalar (`TAKSİT PLANI — dönem bazında...`) asla dar hücreye hapsedilip 4-6 satıra bölünemez (`wrap_text=False`).
- Başlıklar sol kenar payıyla (Row 1: 34pt yükseklik, Montserrat 15pt Bold) saf zemin üzerinde zarif bir kurumsal antet olarak yer alır.
- Açıklama satırı (Row 3: 20pt yükseklik, Segoe UI 9.5pt Italic) tek satırda rahatça okunur.

## 249.2 Terzi Usulü Kolon Genişliği ve Dikey Ritim (Bespoke Proportions & Vertical Rhythm)
- Rastgele sütun genişliği yasaktır:
  - Kod / ID Kolonları: 13.0 - 14.0 pt (Ortalı)
  - Banka / Firma / Cari Adları: 24.0 - 28.0 pt (Sola dayalı, ferah)
  - Açıklama ve Kredi / Ürün Türleri: 22.0 - 32.0 pt
  - Finansal Tutar / Para Kolonları: 18.0 - 22.0 pt (Sağa dayalı, ₺#,##0 formatında)
  - Tarihler ve Durum Rozetleri: 14.0 - 18.0 pt (Ortalı)
- Dikey Ritim Standartları:
  - Satır 1 (Sayfa Başlığı): 34.0 pt
  - Satır 2 (Boşluk/Nefes): 6.0 pt
  - Satır 3 (Alt Açıklama): 20.0 pt
  - Satır 4 (Boşluk): 10.0 pt
  - Satır 5 (Tablo Başlıkları): 28.0 pt (Slate 800 #1E293B dolgu, beyaz kalın yazı)
  - Satır 6–35 (Zebra Tablo Satırları): Düzenli ve standart 22.0 pt

## 249.3 5 Aşamalı Mantıksal CFO İş Akışı Sıralaması (Process Ordering)
Her çalışma kitabında sekmeler kullanıcının zihnini yormayan, işi yapış mantığına göre sıralanır:
1. `AŞAMA 1: GİRİŞ & YOL HARİTASI` -> KILAVUZ (Zümrüt Yeşil - #059669)
2. `AŞAMA 2: STRATEJİK YÖNETİM KOKPİTİ` -> PANO (Koyu Lacivert - #1E293B)
3. `AŞAMA 3: GÜNLÜK İŞLEM & VERİ HAVUZU` -> GİRDİLER (Kehribar Sarı - #D97706), BANKALAR, LIMITLER, KREDILER, TAKSIT_PLANI, ODEME_TAKVIMI (Açık Mavi - #60A5FA)
4. `AŞAMA 4: ARBİTRAJ, STRES TESTİ VE RAPORLAMA` -> REFINANSMAN, YÖNETİM_RAPORU, RAPOR (Mor - #7C3AED)
5. `AŞAMA 5: ARKA ALGORİTMA, MOTOR VE YÖNETİŞİM` -> KARAR_MOTORU, MOTOR, SENARYO_DUYARLILIK, AYARLAR, LISTELER (Koyu Gri - #475569) + 00_BASLANGIC...05_KARAR_SATIS

## 249.4 Referans Kart Mimarisi (KILAVUZ 1:1 Geometrisi)
- KILAVUZ sayfası 3 geniş simetrik karttan (her biri 43.8 pt genişliğinde [B..F], [H..L], [N..R]) oluşur.
- Dış çerçeve: #F2F2F2 platin gri; iç metin alanı saf beyaz (#FFFFFF); butonlar #1B365D kurumsal lacivert.
- Kartların altında mutlaka doğrudan ilgili sekmelere zıplayan (`#PANO!A1`, `#GİRDİLER!A1`, vb.) 5 Aşamalı Proses Tablosu yer alır.

## 249.5 $1,000 Değer Kriterleri (CFO Karar Gücü)
Şablon yalnızca veri kaydetmez; şu 3 stratejik kararı canlı üretir:
1. `Efektif Faiz & Gizli Masraf Tespiti (XIRR):` Dosya masrafları ve komisyonlar dahil gerçek maliyet.
2. `Refinansman Arbitrajı:` Krediyi başka bankaya taşıma durumunda erken kapama cezası düşülmüş net TL kazancı.
3. `Likidite Stres Radarı:` -%20 nakit daralmasında 30/60/90 gün temerrüt ve nakit açığı erken uyarı sinyali.
"""

if "# 249 — GOLDEN RATIO UI" not in content:
    # Insert right before "# 248 — $1,000 SORUMLULUK CÜMLESİ" or right before the closing """
    target_pos = content.find('# 248 — $1,000 SORUMLULUK CÜMLESİ')
    if target_pos != -1:
        updated_content = content[:target_pos] + ADDITIONAL_MANDATE_RULES + "\n\n" + content[target_pos:]
        kernel_path.write_text(updated_content, encoding="utf-8")
        print("✓ KERNEL.py updated with Mandate Rule 249 (Golden Ratio & CFO Workflow Protocol).")
    else:
        print("Target position not found in KERNEL.py")
else:
    print("MANDATE Rule 249 already present in KERNEL.py.")

# =============================================================================
# 2. UPDATE EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx
# =============================================================================
master_xlsx_path = Path("/Users/macair1/projects/excel-arsiv/calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx")
wb_master = openpyxl.load_workbook(master_xlsx_path, data_only=False)

# Update 03_GORSEL_TASARIM sheet
if "03_GORSEL_TASARIM" in wb_master.sheetnames:
    ws_design = wb_master["03_GORSEL_TASARIM"]
    new_rules = [
        ("Altın oran dikey ritim", "Satır yükseklikleri matematiksel hiyerarşide sabitlenir", "Başlık: 34pt, Alt başlık: 20pt, Tablo başlığı: 28pt, Veri: 22pt", "Rastgele satır yükseklikleri"),
        ("Sıfır metin kırılması", "Açıklama ve başlıklar tek satırda akıcıdır (wrap_text=False)", "Geniş sütun ve tek satır başlık; hücre içine sıkışan metin yasaktır", "4-6 satıra bölünen dar metinler"),
        ("5 Aşamalı CFO proses sırası", "Sekmeler iş akışına göre fiziksel olarak dizilir", "01.KILAVUZ -> 02.PANO -> 03.GİRDİLER -> 04-11.OPERASYON -> 12-14.KARAR/RAPOR -> 15-20.MOTOR", "Dağınık alfabetik sheet dizilimi"),
        ("Referans KILAVUZ mimarisi", "1:1 Simetrik 3 geniş kart + 5 Aşamalı Proses Tablosu", "43.8 pt genişliğinde 3 kart, #F2F2F2 çerçeve, #1B365D butonlar ve interaktif köprüler", "Dar, sıkışık veya dekoratif kılavuz"),
        ("Terzi usulü kolon genişlikleri", "Her veri tipi için özel tanımlı ferah genişlik", "ID: 14, İsim: 26, Açıklama: 28, Finansal Tutar: 20, Tarih: 16", "Varsayılan 8.43 genişliği / dar kolonlar")
    ]
    start_row = 23
    for r_i, (k, b, c, d) in enumerate(new_rules, start=start_row):
        ws_design.cell(r_i, 1, k).font = Font(name="Segoe UI", size=9.5, bold=True)
        ws_design.cell(r_i, 2, b).font = Font(name="Segoe UI", size=9.5)
        ws_design.cell(r_i, 3, c).font = Font(name="Segoe UI", size=9.5)
        ws_design.cell(r_i, 4, d).font = Font(name="Segoe UI", size=9.5)
        for col_idx in range(1, 5):
            ws_design.cell(r_i, col_idx).border = Border(
                left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"),
                top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1")
            )
    print("✓ 03_GORSEL_TASARIM updated with Golden Ratio & Zero-Wrap standards.")

# Update 02_KONTROL_LISTESI
if "02_KONTROL_LISTESI" in wb_master.sheetnames:
    ws_ctrl = wb_master["02_KONTROL_LISTESI"]
    new_controls = [
        ("K22", "Metin kırılması (wrap_text) ve dar hücreye sıkışma var mı?", "Hayır, sıfır metin kırılması", "EVET"),
        ("K23", "Sekmeler 5 Aşamalı CFO iş akışına göre dizilmiş mi?", "Evet (KILAVUZ->PANO->GİRDİLER->OPERASYON->RAPOR->MOTOR)", "EVET"),
        ("K24", "KILAVUZ 1:1 referans geniş kart mimarisine ve butonlara sahip mi?", "Evet (43.8 pt kartlar + doğrudan köprüler)", "EVET"),
        ("K25", "Refinansman net nakit kazancı ve -%20 stres testi canlı bağlı mı?", "Evet, formüller PANO ve RAPOR ile entegre", "EVET")
    ]
    start_r = 25
    for r_i, (k, b, c, d) in enumerate(new_controls, start=start_r):
        ws_ctrl.cell(r_i, 1, k).font = Font(name="Segoe UI", size=9.5, bold=True)
        ws_ctrl.cell(r_i, 2, b).font = Font(name="Segoe UI", size=9.5)
        ws_ctrl.cell(r_i, 3, c).font = Font(name="Segoe UI", size=9.5)
        ws_ctrl.cell(r_i, 4, d).font = Font(name="Segoe UI", size=9.5, bold=True, color="047857")
        for col_idx in range(1, 5):
            ws_ctrl.cell(r_i, col_idx).border = Border(
                left=Side(style='thin', color="CBD5E1"), right=Side(style='thin', color="CBD5E1"),
                top=Side(style='thin', color="CBD5E1"), bottom=Side(style='thin', color="CBD5E1")
            )
    print("✓ 02_KONTROL_LISTESI updated with K22-K25 quality gates.")

wb_master.save(master_xlsx_path)
print("✓ EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx saved successfully.")
