from pathlib import Path

path = Path('src/pages/ozel-excel-sistemleri.astro')
text = path.read_text(encoding='utf-8')
canonical = '  <link rel="canonical" href="https://excelarsiv.com/ozel-excel-sistemleri">\n'
block = '''  <meta name="description" content="Mali müşavirler, muhasebe ve finans ekipleri ile KOBİ'ler için işleyişe özel Excel tabloları, finansal otomasyon, nakit akışı, raporlama ve kontrol sistemleri.">\n  <meta name="robots" content="index, follow, max-image-preview:large, max-snippet:-1, max-video-preview:-1">\n  <script type="application/ld+json">\n    {"@context":"https://schema.org","@type":"Service","name":"İşletmeye Özel Excel Sistemleri ve Finansal Otomasyon","url":"https://excelarsiv.com/ozel-excel-sistemleri","description":"Mali müşavirler, muhasebe ve finans ekipleri ile KOBİ'ler için işleyişe özel Excel tabloları, finansal otomasyon, raporlama ve kontrol sistemleri.","provider":{"@type":"Person","name":"Barış Bağırlar","url":"https://excelarsiv.com/hakkinda"},"areaServed":{"@type":"Country","name":"Türkiye"}}\n  </script>\n'''
if '<meta name="description"' in text and '<meta name="robots"' in text and 'application/ld+json' in text:
    print('SEO metadata already present')
elif canonical in text:
    text = text.replace(canonical, canonical + block, 1)
    path.write_text(text, encoding='utf-8')
    print('SEO metadata added')
else:
    raise SystemExit('canonical marker missing')
