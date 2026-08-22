from pathlib import Path
import re

p = Path('src/pages/ozel-excel-sistemleri.astro')
s = p.read_text(encoding='utf-8')

m = re.search(r'(<section[^>]*id="iletisim"[\s\S]*?<h2[^>]*>)([\s\S]*?)(</h2>)', s)
if not m:
    raise SystemExit('CTA iletisim bolumu veya H2 bulunamadi')
new_heading = 'İhtiyaca Yönelik Excel Tabloları Hazırlarım'
s = s[:m.start(2)] + new_heading + s[m.end(2):]
p.write_text(s, encoding='utf-8')
print('CTA heading updated:', new_heading)
