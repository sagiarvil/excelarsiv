# Production commerce function recovery

Bu dosya iş mantığını değiştirmez. 27.08.2026 production Hosting release kontrolünde `createCheckout`, `verifyShopierOrder`, `recoverPurchase`, `createDownloadToken`, `downloadFile`, `requestProofDemo` ve `downloadProofDemo` endpointlerinin production ortamında bulunmadığı doğrulandı.

Amaç: mevcut production workflow'un `functions/` değişikliğini `backend_changed=true` olarak algılamasını ve test edilmiş Firebase Functions deployment zincirini yeniden çalıştırmasını sağlamak.

Kabul kriterleri:
- gerekli 8 production Function `gcloud functions describe` ile bulunur;
- `/api/checkout` bilinmeyen ürün için HTTP 400 + `UNKNOWN_PRODUCT` döndürür;
- gerçek katalog ürünü için checkout HTTP 201 döndürür;
- Proof Demo üretim ve tek kullanımlık indirme akışı geçer;
- custom domain ve canlı SEO kontratı release sonrasında doğrulanır.
