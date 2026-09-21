# PARAMETRE KAYNAKLARI — POS, KOMİSYON VE NET TAHSİLAT KONTROL SİSTEMİ

AYARLAR sayfasındaki her parametrenin kaynağı ve yürürlük tarihi burada özetlenir. Dosyada `anahtar | deger | birim | alt | ust | aciklama | kaynak | yururluk_tarihi | dogrulama_tarihi` sütunlarıyla saklanır (D10).

| Parametre | Varsayılan | Kaynak |
|---|---|---|
| Firma Unvanı | Örnek Firma A.Ş. | Kullanıcı |
| Rapor Tarihi | 08.08.2026 | Kullanıcı |
| Raporu Hazırlayan | Finans Direktörü | Kullanıcı |
| Hedef Net Tahsilat Oranı | 0,97 | Varsayım (sektör pratiği) |
| İade Oranı Risk Eşiği | 0,05 | Varsayım |
| Komisyon Oranı Risk Eşiği | 0,03 | Varsayım |
| İyimser/Baz/Kötümser Senaryo Çarpanı | 1,05 / 1,00 / 0,95 | Varsayım |
| Tornado Komisyon/İade Etkisi | %1 | Varsayım |
| Tahmin Çarpanı | 1,645 | Varsayım (normal dağılım %90 güven aralığı) |
| Kalite Skoru İyi/Orta Eşiği | 90 / 70 | Varsayım |
| P90 Yüzdelik Oranı | 0,90 | Varsayım |
| Aksiyon 1-3 Açıklamaları | Metin | Varsayım |

**Not:** Bu dosyada vergi, SGK, faiz veya mevzuat parametresi sabitlenmemiştir; tüm eşikler kullanıcı tarafından değiştirilebilir varsayımlardır. Mevzuat parametresi kullanılması hâlinde kaynağa (kurum + tebliğ/karar no + tarih) aynı tabloda yer verilir.

## Mevzuat dayanağı

- Türk Ticaret Kanunu (6102 s.): düzenli ve gerçeğe uygun kayıt ilkeleri.
- Vergi Usul Kanunu (213 s.): belge ve kayıt nizamı.

**Veri gizliliği:** Dosya tamamen çevrimdışıdır; verileriniz cihazınızdan çıkmaz (G20).
