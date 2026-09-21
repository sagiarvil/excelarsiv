# PARAMETRE KAYNAKLARI — GÜNLÜK GELİR-GİDER VE GERÇEK KÂRLILIK SİSTEMİ

Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Aşağıda AYARLAR sayfasındaki her parametrenin kaynağı listelenir.

## Mevzuat kaynakları

| Parametre | Kaynak |
|---|---|
| Gerçeğe uygun ve dönemsel kayıt ilkeleri | Türk Ticaret Kanunu (6102 s.) — ticari defter ve kayıt düzeni |
| Kayıt düzeni, defter ve belge nizamı | Vergi Usul Kanunu (213 s.) — kayıt ve muhafaza hükümleri |
| Tahakkuk esası ve dönem ayrımı | Türkiye Muhasebe Standartları çerçevesinde dönemsellik ilkesi |

## Kullanıcı tarafından girilen parametreler

| Parametre | Birim | Kaynak |
|---|---|---|
| Firma Unvanı | Metin | Kullanıcı |
| Vergi Numarası | Metin | Kullanıcı |
| Rapor Tarihi | Tarih | Kullanıcı |
| Para Birimi | Metin | Kullanıcı |
| Dosya Sürümü | Metin | Kullanıcı |
| Lisans Kodu | Kod | Kullanıcı |
| Rapor Dönemi Başlangıcı | Tarih | Kullanıcı |
| Veri Başlangıç Satırı | Skor | Kullanıcı |
| Dönem Başı Stok | TL | Kullanıcı |
| Dönem Sonu Stok | TL | Kullanıcı |
| Yuvarlama Toleransı | TL | Kullanıcı |
| Kaynak URL | Metin | Kullanıcı |
| Kaynak Tarihi | Tarih | Kullanıcı |
| Rapor Hazırlayan | Metin | Kullanıcı |

## Varsayım etiketli parametreler (kaynak = Varsayım)

Bu parametreler ürünün kurulumunda makul işletme varsayımlarıyla doldurulmuştur; kullanıcı kendi sektörüne göre AYARLAR'dan değiştirmelidir.

| Parametre | Varsayılan değer | Açıklama |
|---|---|---|
| Brüt Marj Hedefi | %35 | Brüt marjın altına inilirse KARAR_MOTORU RİSKLİ üretir |
| Faaliyet Marjı Hedefi | %10 | Sabit gider verimliliğini ölçer |
| Net Kâr Marjı Hedefi | %5 | Hedef kâr ciro hesabında kullanılır |
| Nakit Dönüşüm Hedefi | %80 | Kârın nakde dönüşme oranı hedefi |
| Finansman Gideri Eşiği | %30 | Finansman giderinin faaliyet kârına oranı eşiği |
| Ortak Çekişi Eşiği | %100 | Ortak çekişinin net kâra oranı üst eşiği |
| Vergi Oranı | %20 | Vergi oranı varsayımı; kullanıcı günceller |
| Hata Ağırlığı | 5 puan | Veri kalite skorundan her hata için düşülen puan |
| Uyarı Ağırlığı | 10 puan | Veri kalite skorundan her uyarı için düşülen puan |
| Veri Kalite Yüksek/Orta Eşiği | 80 / 50 | Güven düzeyi sınırları |
| Standart Termin | 15 gün | KARAR_MOTORU aksiyonları için standart termin |
| Uzun Termin | 20 gün | Yüksek öncelikli aksiyonlar için uzun termin |
| Marj Bant Oranı | %10 | Brüt/faaliyet marjında NORMAL bandı üst sınırı |
| Nakit/Kâr Asgari Oranı | %50 | Kârın bu oranının altında nakde dönüşmesi RİSKLİ sayılır |
| Olağandışı Gelir Eşiği | %20 | Olağandışı net gelirin net sonuca oranı eşiği |

## Not

Vergi, SGK ve faiz oranları sabitlenmemiştir; güncel mevzuata göre kullanıcı tarafından güncellenmelidir. Hesaplar mali danışmanlık görüşünün yerine geçmez; karar destek amaçlıdır.
