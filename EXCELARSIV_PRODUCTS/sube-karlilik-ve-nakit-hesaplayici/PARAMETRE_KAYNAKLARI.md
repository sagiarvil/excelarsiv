# PARAMETRE KAYNAKLARI — ŞUBE KÂRLILIK VE NAKİT HESAPLAYICI

Aşağıdaki tablo, dosyadaki AYARLAR sayfasında bulunan tüm parametrelerin varsayılan değerlerini ve kaynaklarını listeler. Mevzuat kaynaklı parametreler kurum + tebliğ/karar no + tarih biçimiyle verilir; kaynağı olmayan değerler `Varsayım` olarak işaretlenir.

| Parametre | Varsayılan | Birim | Açıklama | Kaynak |
|---|---|---|---|---|
| Firma unvanı | Örnek Perakende A.Ş. | Metin | Rapor başlığında görünür | Kullanıcı |
| Rapor tarihi | 10.08.2026 | Tarih | Raporun dayanak tarihi | Kullanıcı |
| Raporu hazırlayan | Mali İşler ve Operasyon Yöneticisi | Metin | Rapor künyesi | Kullanıcı |
| Dosya sürümü | 1.0.0 | Metin | Sürüm damgası | Kullanıcı |
| Hedef kâr marjı eşiği | %15 | Oran | KARLI/İNCELE karar eşiği | Varsayım (yönetim hedefi) |
| Kalite eşikleri (İyi/Orta) | 90 / 70 | Puan | Veri kalite skoru sınıfları | Varsayım |
| Senaryo hasılat çarpanları (İyi/Baz/Kötü/Kritik) | 0,90 / 1,00 / 1,15 / 1,30 | Oran | Nakit senaryo bandı | Varsayım (belirsizlik bandı) |
| Tahmin güven çarpanı | 1,645 | Çarpan | %90 güven bandı (normal dağılım) | Varsayım (istatistik standart) |
| Tahmin dönemi | 5 | Dönem | Gelecek dönem tahmin noktası | Varsayım |
| Tornado adımları (Gider/Ciro) | %10 / %5 | Oran | Duyarlılık analizi adım genişliği | Varsayım |
| P90 oranı | %90 | Oran | Kayıt tutarı üst yüzdelik dilimi | Varsayım |
| Yüksek gider eşiği | 100.000 | TL | Yüksek gider tutarı kontrol eşiği | Varsayım |
| Yeterli veri kayıt sayısı | 25 | Adet | Kalite skorunun %100 sayılması için kayıt sayısı | Varsayım |
| Aksiyon önerileri (3 adet) | Metin | Metin | Karar kapısına göre gösterilen öneriler | Varsayım (yönetim kuralı) |
| Panel sıra numaraları (1-6) | 1..6 | Satır | PANO grafiklerinin şube sırası | Kullanıcı |

## Mevzuat dayanağı
Bu ürün mevzuattan bağımsız bir karar destek aracıdır; doğrudan vergi kanunu maddesine bağlı hesaplama içermez. Şube kârlılığı, kâr marjı ve nakit pozisyonu hesapları genel muhasebe ve nakit akışı ilkelerine dayanır; marj eşikleri ve senaryo çarpanları işletme bazında kullanıcı tarafından güncellenebilir. Veriler, Türkiye Finansal Raporlama Standardı'na göre dönem sonucunu etkilemez; yalnızca yönetim kararı desteği sağlar.

## Veri gizliliği
Dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz. Makro, dış bağlantı ve telemetri içermez.
