from pathlib import Path

path = Path('src/pages/ozel-excel-sistemleri.astro')
text = path.read_text(encoding='utf-8')
needle = '  <title>Özel Excel Sistemleri & Finansal Otomasyon | Barış Bağırlar</title>\n'
canonical = '  <link rel="canonical" href="https://excelarsiv.com/ozel-excel-sistemleri">\n'
if canonical in text:
    print('canonical already present')
elif needle in text:
    text = text.replace(needle, needle + canonical, 1)
    path.write_text(text, encoding='utf-8')
    print('canonical added')
else:
    raise SystemExit('title marker missing')
