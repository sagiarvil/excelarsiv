import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

master_path = 'calisma_prorokolu/EXCELARSIV_DECISION_OS_V15_1_MASTER_STANDART_UI_LOCKED.xlsx'
wb = openpyxl.load_workbook(master_path)

data_sheets = [
    '00_BASLANGIC', '01_YURUTME_CEKIRDEGI', '02_KONTROL_LISTESI', 
    '03_GORSEL_TASARIM', '04_MUHASEBE_THP', '05_KARAR_SATIS'
]

for s in data_sheets:
    if s in wb.sheetnames:
        ws = wb[s]
        ws.views.sheetView[0].showGridLines = False
        # find table max col
        table_cols = [c for c in range(1, 20) if ws.cell(4, c).value is not None or ws.cell(5, c).value is not None]
        max_c = max(table_cols) if table_cols else 8
        col_let = openpyxl.utils.get_column_letter(max_c)
        for mr in list(ws.merged_cells.ranges):
            if 'A1' in mr or 'B1' in mr or 'A2' in mr or 'B2' in mr:
                ws.merged_cells.remove(mr)
        if ws['A1'].value or ws['B1'].value:
            start_col = 'B' if ws['B1'].value and not ws['A1'].value else 'A'
            ws.merge_cells(f'{start_col}1:{col_let}1')
            ws[f'{start_col}1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

# Kural ekleme: 03_GORSEL_TASARIM
if '03_GORSEL_TASARIM' in wb.sheetnames:
    ws_v = wb['03_GORSEL_TASARIM']
    next_r = ws_v.max_row + 1
    ws_v[f'B{next_r}'] = "Sıfır Metin Kırılması (Zero Subtitle Wrap)"
    ws_v[f'C{next_r}'] = "Tüm sayfalarda A1 (Banner) ve A3 (Açıklama/Alt Başlık) hücreleri tablo genişliği kadar (A1:X1 / A3:X3) birleştirilmeli ve wrap_text=False olmalıdır. Metinlerin tek sütuna sıkışıp dikey kırılması kesinlikle yasaktır."
    ws_v[f'D{next_r}'] = "ZORUNLU"

wb.save(master_path)
print("Master Template: Sıfır metin kırılması kuralı ve düzenlemesi başarıyla tamamlandı.")
