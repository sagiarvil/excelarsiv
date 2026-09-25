import glob, openpyxl, re

files = glob.glob('EXCELARSIV_PRODUCTS/**/*.xlsx', recursive=True)
files = [f for f in files if not f.split('/')[-1].startswith('~$')]

# openpyxl OpenXML evrensel fonksiyon haritası
# Formüllerde Türkçe adlar ve noktalı virgül (;) Excel OpenXML standardında #NAME? / #AD? üretir.
# OpenXML formatında tüm fonksiyonlar İNGİLİZCE ve parametre ayırıcı VİRGÜL (,) olmalıdır.

def clean_formula_syntax(formula_str):
    if not isinstance(formula_str, str) or not formula_str.startswith('='):
        return formula_str

    f = formula_str

    # 1. Türkçe Fonksiyon Adlarını İngilizce Karşılıklarına Çevir
    replacements = [
        (r'\bTOPLA\b', 'SUM'),
        (r'\bEĞER\b', 'IF'),
        (r'\bDÜŞEYARA\b', 'VLOOKUP'),
        (r'\bYATAYARA\b', 'HLOOKUP'),
        (r'\bORTALAMA\b', 'AVERAGE'),
        (r'\bBAĞ_DEĞ_DOLU_SAY\b', 'COUNTA'),
        (r'\bBAĞ_DEĞ_SAY\b', 'COUNT'),
        (r'\bBOŞLUKSAY\b', 'COUNTBLANK'),
        (r'\bÇOKETOPLA\b', 'SUMIFS'),
        (r'\bETOPLA\b', 'SUMIF'),
        (r'\bÇOKEĞERSAY\b', 'COUNTIFS'),
        (r'\bEĞERSAY\b', 'COUNTIF'),
        (r'\bEĞERHATA\b', 'IFERROR'),
        (r'\bYUVARLA\b', 'ROUND'),
        (r'\bAŞAĞIYUVARLA\b', 'ROUNDDOWN'),
        (r'\bYUKARIYUVARLA\b', 'ROUNDUP'),
        (r'\bBUGÜN\b', 'TODAY'),
        (r'\bŞİMDİ\b', 'NOW'),
        (r'\bSOLDAN\b', 'LEFT'),
        (r'\bSAĞDAN\b', 'RIGHT'),
        (r'\bPARÇAAL\b', 'MID'),
        (r'\bBÜYÜKHARF\b', 'UPPER'),
        (r'\bKÜÇÜKHARF\b', 'LOWER'),
        (r'\bMUTLAK\b', 'ABS'),
        (r'\bKAREKÖK\b', 'SQRT'),
        (r'\bVE\b', 'AND'),
        (r'\bYADA\b', 'OR')
    ]

    for tr_pat, en_name in replacements:
        f = re.sub(tr_pat, en_name, f, flags=re.IGNORECASE)

    # 2. Noktalı virgülleri (;) tırnak içi metinlere dokunmadan virgül (,) ile değiştir
    # Metin kısımlarını (string literal "...") korumak için regex:
    parts = re.split(r'("[^"]*")', f)
    for i in range(len(parts)):
        # Çift indeksler formül kodudur (tırnak dışı)
        if i % 2 == 0:
            parts[i] = parts[i].replace(';', ',')

    f_fixed = ''.join(parts)
    return f_fixed

print(f"Toplam {len(files)} dosya için Evrensel Formül İyileştirici Başlatılıyor...")
total_fixed_cells = 0
fixed_files_count = 0

for idx, fpath in enumerate(files, start=1):
    try:
        wb = openpyxl.load_workbook(fpath, data_only=False)
        wb_modified = False
        
        for sname in wb.sheetnames:
            ws = wb[sname]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    val = cell.value
                    if val and isinstance(val, str) and val.startswith('='):
                        fixed = clean_formula_syntax(val)
                        if fixed != val:
                            cell.value = fixed
                            total_fixed_cells += 1
                            wb_modified = True
                            
        if wb_modified:
            wb.save(fpath)
            fixed_files_count += 1
    except Exception as e:
        print(f"HATA: {fpath} -> {e}")

print(f"\nİŞLEM TAMAMLANDI:")
print(f"Düzeltilen Dosya Sayısı: {fixed_files_count}")
print(f"Düzeltilen Toplam Formül Sayısı: {total_fixed_cells}")
