import glob, openpyxl

files = glob.glob('EXCELARSIV_PRODUCTS/**/*.xlsx', recursive=True)
files = [f for f in files if not f.split('/')[-1].startswith('~$')]

turkish_funcs = [
    'TOPLA(', 'EĞER(', 'DÜŞEYARA(', 'YATAYARA(', 'ORTALAMA(', 
    'BAĞ_DEĞ_DOLU_SAY(', 'ÇOKETOPLA(', 'EĞERSAY(', 'YUVARLA(', 
    'BUGÜN(', 'ŞİMDİ(', 'BOŞLUKSAY(', 'SOLDAN(', 'SAĞDAN(', 
    'PARÇAAL(', 'BÜYÜKHARF(', 'KÜÇÜKHARF(', 'MUTLAK(', 'KAREKÖK('
]

remaining_defects = []

for f in files:
    try:
        wb = openpyxl.load_workbook(f, data_only=False)
        for sname in wb.sheetnames:
            ws = wb[sname]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    val = cell.value
                    if val and isinstance(val, str) and val.startswith('='):
                        val_up = val.upper()
                        for tf in turkish_funcs:
                            if tf in val_up:
                                remaining_defects.append((f, sname, cell.coordinate, val))
    except Exception as e:
        pass

print(f"Denetlenen Dosya Sayısı: {len(files)}")
print(f"Kalan Türkçe Fonksiyon Hatası: {len(remaining_defects)}")
for d in remaining_defects[:5]:
    print("  Kusur:", d)
