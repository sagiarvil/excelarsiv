# PARAMETRE KAYNAKLARI — ÜRETİM REÇETESİ & ZAM YANSITMA HESAPLAYICI

Aşağıdaki tablo, dosyadaki AYARLAR sayfasında bulunan tüm parametrelerin varsayılan değerlerini ve kaynaklarını listeler. Mevzuat kaynaklı parametreler kurum + tebliğ/karar no + tarih biçimiyle verilir; kaynağı olmayan değerler `Varsayım` olarak işaretlenir.

| Parametre | Varsayılan | Birim | Açıklama | Kaynak |
|---|---|---|---|---|
| Firma unvanı | Örnek Firma A.Ş. | Metin | Rapor başlığında görünür | Kullanıcı |
| Rapor tarihi | 10.08.2026 | Tarih | Raporun dayanak tarihi | Kullanıcı |
| Raporu hazırlayan | Üretim ve Mali İşler Müdürü | Metin | Rapor künyesi | Kullanıcı |
| Dosya sürümü | 1.0.0 | Metin | Sürüm damgası | Kullanıcı |
| Hedef kâr marjı | %30 | Oran | Önerilen satış fiyatı bu marjı korur | Varsayım (işletme politikası) |
| Zam yansıtma eşiği | %3 | Oran | Üzerindeki genel artışta ZAM YANSIT kararı | Varsayım (yönetim eşiği) |
| Kritik fire oranı | %10 | Oran | Üzerindeki fire oranı kritik kabul edilir | Varsayım (sektör pratiği) |
| İşçilik oranı | %12 | Oran | Hammadde maliyetine eklenir | Varsayım (firma gider yapısı) |
| GÜG oranı | %8 | Oran | Hammadde maliyetine eklenir | Varsayım (firma gider yapısı) |
| Senaryo çarpanları (İyi/Baz/Kötü/Kritik) | 0 / %3 / %8 / %15 | Oran | Zam senaryo bant aralığı | Varsayım (belirsizlik bandı) |
| Tahmin güven çarpanı | 1,645 | Çarpan | %90 güven bandı (normal dağılım) | Varsayım (istatistik standart) |
| Kalite eşikleri (İyi/Orta) | 90 / 70 | Puan | Veri kalite skoru sınıfları | Varsayım |
| Malzeme payı eşiği | %50 | Oran | Tek hammadde yoğunlaşma eşiği | Varsayım |
| P90 oranı | %90 | Oran | Birim fiyat üst yüzdelik dilimi | Varsayım |
| Temel giriş kolonu sayısı | 8 | Adet | Veri kalitesi skor kolonu sayısı | Varsayım |
| Tornado adımları (Un/Fire) | %10 / %5 | Oran | Duyarlılık analizi adım genişliği | Varsayım |
| Yüksek maliyet eşiği | 50.000 | TL | Kontrol eşiği | Varsayım |

## Mevzuat dayanağı
Bu ürün mevzuattan bağımsız bir karar destek aracıdır; doğrudan vergi kanunu maddesine bağlı hesaplama içermez. Reçete maliyeti ve zam yansıtma mantığı genel muhasebe ilkelerine (tam maliyet yöntemi) uygun olarak kurulmuştur; işletme bazında oranlar kullanıcı tarafından güncellenebilir.

## Veri gizliliği
Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Makro, dış bağlantı ve telemetri içermez.
