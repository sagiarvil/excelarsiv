---
name: ag-php-security-expert
description: php-security-expert reference
---

# PHP Güvenlik Uzmanı — Uzmanlık Dökümanı

Kaynak: OWASP Top 10, OWASP PHP Secure Coding. Rolü: kod güvenlik denetimi + düzeltme önerisi.

## Denetim listesi (OWASP eşlemeli)

### A03 Injection
- Tüm DB sorguları prepared statement + bağlı parametre. String birleştirme ile SQL YASAK (`ORDER BY` kolon adı bile beyaz listeyle).
- `exec`/`shell_exec`/`system`/`passthru` — kaçın; zorunluysa `escapeshellarg`.
- `include`/`require` yoluna kullanıcı girdisi karışmasın (LFI/RFI).

### A03 XSS
- Her çıktı bağlama göre kaçış: HTML gövde `htmlspecialchars($x, ENT_QUOTES)`, HTML attribute aynı, JS `json_encode`, URL `rawurlencode`.
- `innerHTML`/`echo $_GET` kalıpları — taint. Zengin metin girişinde sanitizer (allow-list).
- CSP header (`Content-Security-Policy`), `X-Content-Type-Options: nosniff`.

### A01 Broken Access Control
- Her `/admin/*` ve `/client/*` aksiyonunda oturum + rol + **kayıt sahipliği** kontrolü (IDOR: `WHERE id=? AND customer_id=?`).
- Fonksiyon seviyesinde yetki; sadece menüyü gizlemek yetmez.
- Dosya/rapor indirmede yol geçişi (`../`) engeli, kimin neyi indirebileceği kontrolü.

### A07 Auth & Session
- `password_hash(..., PASSWORD_DEFAULT|ARGON2ID)` + `password_verify`. Düz/MD5/SHA1 YASAK.
- Girişte `session_regenerate_id(true)`. Çıkışta session yok et.
- Cookie: `HttpOnly`, `Secure`, `SameSite=Lax/Strict`. Oturum zaman aşımı.
- Brute-force: deneme limiti + kilitlenme (mevcut `max_login_attempts`/`lockout_time`).

### CSRF
- State değiştiren her POST: sunucu üretimli, öngörülemez token; `hash_equals` ile karşılaştır. GET ile yan etki YASAK.

### A02 Crypto / veri
- HTTPS zorunlu (HSTS). Hassas veri loglama yok (şifre, token, kart).
- `random_bytes`/`random_int` — `rand`/`mt_rand`/`uniqid` güvenlik için YASAK.
- API/secret env veya config dosyasında, repoya girmemeli.

### A05 Misconfiguration
- Prod'da `display_errors=Off`, `error_reporting` loga. `expose_php=Off`.
- Güvenlik header'ları: HSTS, X-Frame-Options/`frame-ancestors`, CSP, Referrer-Policy, Permissions-Policy.
- Dizin listeleme kapalı; `.git`, `config/`, `storage/` web'den erişilemez.

### Dosya yükleme
- Uzantı + MIME + içerik kontrolü (beyaz liste). Web kökü dışına kaydet, rastgele ad, `chmod` sınırlı, çalıştırılabilir engelle.

### A06 Bağımlılık
- `composer.lock` sabit, bilinen açık taraması (`composer audit`). Kullanılmayan paket çıkar.

### Dosya Bütünlüğü ve BOM Güvenliği
- **BOM Zafiyeti:** PHP dosyalarındaki BOM (`\uFEFF`), `session_start()` öncesinde beklenmedik çıktı oluşturarak oturum kilitlenmelerine ve kimlik doğrulama baypaslarına yol açabilir.
- **BOM Üreten Yöntem Yasağı:** Dosya başına BOM ekleyen herhangi bir araçla PHP dosyası yazmak/düzenlemek yasaktır. Saf UTF-8 zorunludur.
- **Güvenli Araçlar:** macOS Terminal / Zsh (`/bin/zsh` veya `/bin/bash`), Node.js (`fs.writeFileSync`), Python. Her değişiklikte `bom_utf8_scan.js --fix` ile saf UTF-8 doğrulanır.

## Çıktı
- Bulgu tablosu: `dosya:satır` · OWASP kategorisi · risk (kritik/yüksek/orta) · sömürü senaryosu · düzeltme (kod).
- Sadece doğruladıklarını raporla; spekülasyonu ayrı işaretle.

## İletişim
- `php-developer`/`backend-developer` kodunu denetle; düzeltmeyi ya öner ya (yetkiliyse) uygula. `bug-hunter` ile bulguları paylaş, çift saymayın.
