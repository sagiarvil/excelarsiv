# Parametre Kaynakları — Asgari Kurumlar Vergisi Simülasyon Motoru

| Anahtar | Kaynak | Yürürlük | Doğrulama | Açıklama |
|---|---|---|---|---|
| kv_genel_oran (AKV-001) | KVK md.32 | 01.01.2025 | 10.08.2026 | Genel kurumlar vergisi oranı |
| asgari_taban_matrah (AKV-002) | KVK 32/C | 01.01.2025 | 10.08.2026 | Asgari taban matrah (yapılandırılabilir) |
| asgari_oran (AKV-003) | 7524 s.K. md.36 / KVK 32/C | 01.01.2025 | 10.08.2026 | Asgari KV oranı |
| tamamlayici_bayrak (AKV-004) | KVK 32/C | 01.01.2025 | 10.08.2026 | Tamamlayıcı hesap bayrağı |
| yuvarlama_ondalik (AKV-005) | Uygulama notu | 01.01.2025 | 10.08.2026 | Yuvarlama basamağı |

Ortak parametre deposu: `ortak/parametreler.yaml` (`akv_asgari_oran`, `kv_genel_oran`).

Belirsizlik: istisna kırılım sırası mevzuatta tartışmalıdır; dosya yorum A ve B sonucunu yan yana gösterir.
