# Production paid-product Storage recovery

27.08.2026 production commerce recovery sonrasında 8/8 Firebase Function production ortamında mevcut ve `/api/checkout` bilinmeyen ürün için doğru şekilde HTTP 400 + `UNKNOWN_PRODUCT` döndürmektedir. Gerçek ürün checkout testi ise `PRODUCT_NOT_READY` ile durmuştur.

Bu dosya ürün binary içeriğini veya ödeme mantığını değiştirmez. Amaç, mevcut release workflow'un `delivery/paid-products/` katmanını `delivery_changed=true` olarak algılamasını ve test edilmiş private Storage senkronizasyon zincirini çalıştırmasını sağlamaktır.

Kabul kriterleri:
- satışa açık tüm katalog ürünleri local `delivery/paid-products/.../current.*` kaynağından private Firebase Storage'a senkronize edilir;
- `check-paid-products.mjs --strict --verify-local-parity` tüm satışa açık ürünlerde READY + SHA-256 parity verir;
- gerçek katalog ürünü `/api/checkout` üzerinden HTTP 201 döndürür;
- katalogdaki satışa açık ürünlerin hiçbiri delivery eksikliği nedeniyle bloke değildir;
- production Proof Demo akışı geçer;
- custom domain ve canlı SEO kontratı release sonunda doğrulanır.
