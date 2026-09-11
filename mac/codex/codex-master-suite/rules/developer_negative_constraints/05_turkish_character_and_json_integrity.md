# NEGATIVE CONSTRAINTS: Turkish Character & JSON Data Integrity

> **Severity:** CRITICAL / NON-NEGOTIABLE (P0)  
> **Applies To:** AI Model Prose Generation, PHP, JavaScript/TypeScript, Python, SQL, REST APIs & Webhooks

## 🚫 ABSOLUTE PROHIBITIONS (KESİNLİKLE YASAK OLANLAR)

### 1. Prose & Text Output Prohibitions (Cümle ve Metin Kuralları):
- ❌ **NEVER** replace Turkish letters with ASCII equivalents when generating Turkish text (e.g. NEVER write "Turkce", "Istanbul", "degil", "calisiyor" -> ALWAYS write "Türkçe", "İstanbul", "değil", "çalışıyor").
- ❌ **NEVER** drop or corrupt the specific characters: `ç, Ç, ğ, Ğ, ı, I, i, İ, ö, Ö, ş, Ş, ü, Ü`.
- ❌ **NEVER** output or tolerate Mojibake / ISO-8859-9 / Windows-1254 artifacts: `Ý, Ð, Þ, ý, ð, þ`. If encountered, IMMEDIATELY convert to `İ, Ğ, Ş, ı, ğ, ş`.

### 2. Code Generation Prohibitions (Kod Üretim Kuralları):
- ❌ **NEVER use standard ASCII case conversion on Turkish text:**
  - **In PHP:** `strtoupper($text)` and `strtolower($text)` are **STRICTLY BANNED** on Turkish strings.  
    ✅ **MANDATORY:** Use `mb_convert_case($text, MB_CASE_UPPER, "UTF-8")` or dedicated Turkish case mapping.
  - **In JavaScript:** `str.toUpperCase()` and `str.toLowerCase()` without locale are **STRICTLY BANNED** on Turkish strings.  
    ✅ **MANDATORY:** Use `str.toLocaleUpperCase('tr-TR')` and `str.toLocaleLowerCase('tr-TR')`.
  - **Strict I-Rule:** `i` -> `İ` (UPPER), `ı` -> `I` (UPPER), `İ` -> `i` (LOWER), `I` -> `ı` (LOWER). Standard English converts `i` -> `I` which corrupts Turkish words!

### 3. API & Webhook Header Prohibitions:
- ❌ **NEVER omit `charset=utf-8` in JSON endpoints:**
  - Writing `header('Content-Type: application/json');` without charset is **BANNED**.  
    ✅ **MANDATORY:** Always write `header('Content-Type: application/json; charset=utf-8');`.
  - In requests, always include:  
    `Content-Type: application/json; charset=utf-8`  
    `Accept: application/json; charset=utf-8`

### 4. JSON Serialization Prohibitions:
- ❌ **NEVER serialize JSON without Unicode preservation in PHP:**
  - Raw `json_encode($data)` that escapes Turkish letters into \u sequences unnecessarily or corrupts them is **BANNED**.  
    ✅ **MANDATORY:** Always write:  
    `json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);`

### 5. Database Connection Prohibitions:
- ❌ **NEVER connect to MySQL/PostgreSQL without Turkish UTF-8 collation:**
  - Connecting without setting charset is **BANNED**.  
    ✅ **MANDATORY:** Execute immediately upon connection:  
    `SET NAMES utf8mb4 COLLATE utf8mb4_turkish_ci;`
