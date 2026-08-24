# Native Infographic System

Bu katman raster görsel yerine build sırasında gerçek HTML/CSS üretir. Amaç; görsel anlatım, mobil okunabilirlik, iç bağlantı, erişilebilirlik ve arama motoru/LLM okunabilirliğini aynı anda korumaktır.

## Dağıtım

- Ana sayfa: işletme verisi → Excel sistemi → yönetim kararı.
- Özel Excel Sistemleri: sorun → iş kuralı → karar.
- Rehber: kullanıcı sorusu → rehber → ilgili sistem.
- Tüm ürün sayfaları: koyu iş akışı bölümünde kontrast/readability katmanı.
- 6 yüksek niyetli ürün: ürüne özgü karar akışı infografiği.

## İlkeler

- Ana mesaj görsele hapsedilmez; gerçek HTML metnidir.
- Her infografik tek bir karar sorusunu anlatır.
- Aynı kompozisyon her sayfada tekrar edilmez.
- İç bağlantılar görünür ve bağlamsaldır.
- Mobilde tek kolona düşer.
- Build kapısı, eksik injection veya ürün kapsamını hata olarak durdurur.
