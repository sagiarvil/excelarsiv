# PARAMETRE KAYNAKLARI — NAKLİYE MALİYETİ HESAPLAYICI

Aşağıdaki tablo, dosyadaki AYARLAR sayfasında bulunan tüm parametrelerin varsayılan değerlerini ve kaynaklarını listeler. Mevzuat kaynaklı parametreler kurum + tebliğ/karar no + tarih biçimiyle verilir; kaynağı olmayan değerler `Varsayım` olarak işaretlenir.

| Parametre | Varsayılan | Birim | Açıklama | Kaynak |
|---|---|---|---|---|
| Firma unvanı | Örnek Nakliyat A.Ş. | Metin | Rapor başlığında görünür | Kullanıcı |
| Rapor tarihi | 10.08.2026 | Tarih | Raporun dayanak tarihi | Kullanıcı |
| Raporu hazırlayan | Filo ve Mali İşler Yöneticisi | Metin | Rapor künyesi | Kullanıcı |
| Dosya sürümü | 1.0.0 | Metin | Sürüm damgası | Kullanıcı |
| Yakıt fiyatı | 42,00 | TL/lt | Dizel yakıt fiyatı | Varsayım (piyasa ortalaması) |
| Hedef kâr marjı | %15 | Oran | Kâr marjı karar eşiği | Varsayım (filonun kârlılık hedefi) |
| Kalite eşikleri (İyi/Orta) | 90 / 70 | Puan | Veri kalite skoru sınıfları | Varsayım |
| Senaryo yakıt çarpanları (İyi/Baz/Kötü/Kritik) | 0,90 / 1,00 / 1,15 / 1,30 | Oran | Yakıt fiyatı bant aralığı | Varsayım (belirsizlik bandı) |
| Tahmin güven çarpanı | 1,645 | Çarpan | %90 güven bandı (normal dağılım) | Varsayım (istatistik standart) |
| Tahmin dönemi | 5 | Dönem | Gelecek dönem tahmin noktası | Varsayım |
| Tornado adımları (Yakıt/Bakım) | %10 / %5 | Oran | Duyarlılık analizi adım genişliği | Varsayım |
| P90 oranı | %90 | Oran | Birim maliyet üst yüzdelik dilimi | Varsayım |
| Yüksek sefer maliyeti eşiği | 100.000 | TL | Yüksek sefer maliyeti kontrol eşiği | Varsayım |
| Aksiyon önerileri (3 adet) | Metin | Metin | Karar kapısına göre gösterilen öneriler | Varsayım (yönetim kuralı) |
| Panel sıra numaraları (1-6) | 1..6 | Satır | PANO grafiklerinin araç sırası | Kullanıcı |

## Mevzuat dayanağı

Bu ürün mevzuattan bağımsız bir karar destek aracıdır; doğrudan vergi kanunu maddesine bağlı hesaplama içermez. Nakliye maliyeti; yakıt, bakım, sürücü ve amortisman kalemlerinin sefer bazında toplanması ve hasılatla karşılaştırılması mantığına dayanır. Yakıt tüketiminin (lt/100 km) ve bakım oranının (TL/km) araç tipine göre tanımlanması, filo yönetiminde genel kabul gören maliyet muhasebesi ilkeleriyle uyumludur. Yakıt fiyatı, hedef marj ve eşik değerler kullanıcı tarafından güncellenebilir.

## Veri gizliliği

Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Makro, dış bağlantı ve telemetri içermez.
