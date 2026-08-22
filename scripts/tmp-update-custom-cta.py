from pathlib import Path

p = Path('src/pages/ozel-excel-sistemleri.astro')
s = p.read_text(encoding='utf-8')
repls = {
    'Dosyanı İlet / Çözüm Al': 'İhtiyacıma Özel Excel Tablosu',
    'Dosyanızı İletin / 24 Saatte Analiz Alın': 'İhtiyacıma Özel Excel Tablosu Hazırlat',
    '<span>Bu Kaybı Sıfırlayalım (WhatsApp)</span>': '<span>İhtiyacıma Özel Excel Tablosu Hazırlat</span>',
    'Mevcut Excel tablonuzun ekran görüntüsünü veya otomatikleşmesini istediğiniz süreci yazın; <strong>aynı gün içinde nasıl çözüleceğini ve net maliyetini</strong> çıkarıp döneyim.': 'Hazır bir şablona uymak zorunda değilsiniz. İhtiyacınızı, mevcut Excel tablonuzu veya otomatikleşmesini istediğiniz süreci paylaşın; <strong>işleyişinize ve ihtiyacınıza yönelik Excel tablosunu</strong> nasıl hazırlayacağımızı ve kapsamını netleştirip döneyim.',
    "WhatsApp'tan Dosya / Ekran Görüntüsü Gönder": 'İhtiyaca Özel Excel Tablosu Talebi Gönder',
}
for old, new in repls.items():
    if old not in s:
        raise SystemExit(f'Beklenen metin bulunamadı: {old}')
    s = s.replace(old, new)
p.write_text(s, encoding='utf-8')
# workflow trigger
