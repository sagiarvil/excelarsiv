# PARAMETRE KAYNAKLARI — STOK, SATIŞ VE NAKİT BAĞLANMA SİSTEMİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Dosya Sürümü | 1.0.0 | Kullanıcı |
| Devir Hızı Hedef Eşiği | 6 | Varsayım |
| Devir Hızı Risk Eşiği | 3 | Varsayım |
| Stoğa Bağlanan Nakit Eşiği | 500.000 ₺ | Varsayım |
| Ortalama Birim Maliyet Eşiği | 1.000 ₺ | Varsayım |
| İyimser/Baz/Kötümser Senaryo Çarpanı | 1,05 / 1,00 / 0,95 | Varsayım |
| Tornado Fiyat/Adet Etkisi | %1 / %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |
| Panel Satırı 1-12 | 1-12 | Kullanıcı |

**Not:** Bu dosyada vergi, SGK, faiz veya mevzuat parametresi sabitlenmemiştir; tüm eşikler kullanıcı tarafından değiştirilebilir varsayımlardır. Mevzuat parametresi kullanılması hâlinde kaynağa (kurum + tebliğ/karar no + tarih) aynı tabloda yer verilir.

## Mevzuat dayanağı

- Türk Ticaret Kanunu (6102 s.): düzenli ve gerçeğe uygun kayıt ilkeleri.
- Vergi Usul Kanunu (213 s.): envanter ve değerleme esasları.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
