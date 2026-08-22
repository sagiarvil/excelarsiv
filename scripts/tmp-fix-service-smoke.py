from pathlib import Path

path = Path('src/pages/ozel-excel-sistemleri.astro')
text = path.read_text(encoding='utf-8')

text = text.replace('https://wa.me/905393333303?text=', 'https://api.whatsapp.com/send?phone=905393333303&text=')

hero = '  <!-- HERO SECTION -->\n'
footer = '  <!-- FOOTER -->\n'
if '<main>' not in text:
    if hero not in text or footer not in text:
        raise SystemExit('main wrapper markers missing')
    text = text.replace(hero, '  <main>\n\n' + hero, 1)
    text = text.replace(footer, '  </main>\n\n' + footer, 1)

if 'wa.me' in text:
    raise SystemExit('wa.me remains')
if '<main>' not in text or '</main>' not in text:
    raise SystemExit('main wrapper missing')
if 'api.whatsapp.com/send?phone=905393333303' not in text:
    raise SystemExit('WhatsApp contact target missing')

path.write_text(text, encoding='utf-8')
print('service smoke contract patch applied')
