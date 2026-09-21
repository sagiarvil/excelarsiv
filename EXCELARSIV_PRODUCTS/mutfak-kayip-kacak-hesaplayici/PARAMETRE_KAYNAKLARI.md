# PARAMETRE KAYNAKLARI — MUTFAK KAYIP/KAÇAK HESAPLAYICI

Aşağıdaki tablo, dosyadaki AYARLAR sayfasında bulunan tüm parametrelerin varsayılan değerlerini ve kaynaklarını listeler. Mevzuat kaynaklı parametreler kurum + tebliğ/karar no + tarih biçimiyle verilir; kaynağı olmayan değerler `Varsayım` olarak işaretlenir.

| Parametre | Varsayılan | Birim | Açıklama | Kaynak |
|---|---|---|---|---|
| Firma unvanı | Örnek Lokanta A.Ş. | Metin | Rapor başlığında görünür | Kullanıcı |
| Rapor tarihi | 10.08.2026 | Tarih | Raporun dayanak tarihi | Kullanıcı |
| Raporu hazırlayan | Mutfak ve Mali İşler Yöneticisi | Metin | Rapor künyesi | Kullanıcı |
| Dosya sürümü | 1.0.0 | Metin | Sürüm damgası | Kullanıcı |
| İnceleme eşiği (kayıp oranı) | %3 | Oran | Toplam kayıp oranının inceleme eşiği | Varsayım (sektör pratiği) |
| Kritik eşik (kayıp oranı) | %8 | Oran | Toplam kayıp oranının kritik eşiği | Varsayım (sektör pratiği) |
| Kalite eşikleri (İyi/Orta) | 90 / 70 | Puan | Veri kalite skoru sınıfları | Varsayım |
| Senaryo kayıp çarpanları (İyi/Baz/Kötü/Kritik) | 0,5 / 1,0 / 1,5 / 2,0 | Oran | Kayıp bant aralığı | Varsayım (belirsizlik bandı) |
| Tahmin güven çarpanı | 1,645 | Çarpan | %90 güven bandı (normal dağılım) | Varsayım (istatistik standart) |
| Tahmin dönemi | 5 | Dönem | Gelecek dönem tahmin noktası | Varsayım |
| Tornado adımları (Et/Fire) | %10 / %5 | Oran | Duyarlılık analizi adım genişliği | Varsayım |
| P90 oranı | %90 | Oran | Birim maliyet üst yüzdelik dilimi | Varsayım |
| Yüksek kayıp eşiği | 50.000 | TL | Yüksek kayıp tutarı kontrol eşiği | Varsayım |
| Aksiyon önerileri (3 adet) | Metin | Metin | Karar kapısına göre gösterilen öneriler | Varsayım (yönetim kuralı) |
| Panel sıra numaraları (1-6) | 1..6 | Satır | PANO grafiklerinin ürün sırası | Kullanıcı |

## Mevzuat dayanağı
Bu ürün mevzuattan bağımsız bir karar destek aracıdır; doğrudan vergi kanunu maddesine bağlı hesaplama içermez. Kayıp/kaçak tespiti, stok mutabakatı (dönem başı stok + alım − dönem sonu stok) ve reçete bazlı teorik tüketim mantığına dayanır; bu yöntem genel muhasebe ve stok kontrol ilkeleriyle uyumludur. Gıda sektöründe fire/kayıp oranı eşikleri işletme bazında kullanıcı tarafından güncellenebilir.

## Veri gizliliği
Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Makro, dış bağlantı ve telemetri içermez.
