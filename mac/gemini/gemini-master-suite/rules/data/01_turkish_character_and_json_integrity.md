# RULE: Turkish Character & JSON Data Integrity Enforcer

> **Scope:** Universal API, Webhook, Database, HTTP Headers & Data Serialization Standard  
> **Authority:** P0 Constitutional Data Integrity Protocol  
> **Target Platforms:** Node.js, PHP 8.3/8.4, Python 3, Databases (MySQL/PostgreSQL)

---

## 1. Objective
Ensure that all incoming and outgoing data, particularly from `utf8_turkish_ci` or `utf8mb4_turkish_ci` localized databases, is processed, formatted, and transmitted without any character corruption (mojibake). Strict adherence to UTF-8 encoding standards is mandatory for all API requests, responses, webhooks, and inter-service payloads.

---

## 2. Mandatory HTTP Header Rules
Every HTTP request, webhook, and API payload generated or handled by this system MUST strictly enforce and explicitly include the following headers:
- `Content-Type: application/json; charset=utf-8`
- `Accept: application/json; charset=utf-8`

Never omit the `charset=utf-8` parameter in JSON endpoints, webhooks, or API requests.

---

## 3. Turkish Character Handling & Collation Rules
- **No Character Distortion:** Never replace, drop, strip, or corrupt Turkish-specific characters under any circumstances:  
  `ç`, `Ç`, `ğ`, `Ğ`, `ı`, `I`, `i`, `İ`, `ö`, `Ö`, `ş`, `Ş`, `ü`, `Ü`.
- **The Turkish "I" Case Sensitivity Rule (STRICT):** 
  - When converting text to UPPERCASE: `i` MUST become `İ`, and `ı` MUST become `I`.
  - When converting text to lowercase: `İ` MUST become `i`, and `I` MUST become `ı`.
  - **BANNED:** Standard ASCII/English case conversion methods that turn `i` into `I` or `İ` into `I` (e.g. `strtoupper()`, `toLowerCase()`).
  - **In PHP:** Always use `mb_convert_case($str, MB_CASE_UPPER, "UTF-8")` or dedicated Turkish case mapping.
  - **In JavaScript:** Always use `str.toLocaleUpperCase('tr-TR')` and `str.toLocaleLowerCase('tr-TR')`.
  - **In Python:** Use locale-aware mappings for `i` <-> `İ` and `ı` <-> `I`.
- **Database Alignment:**
  - Treat all extracted data with the assumption that it originates from a `utf8_turkish_ci` or `utf8mb4_turkish_ci` schema.
  - Explicitly set the connection encoding to `utf8mb4` immediately after connection (`SET NAMES utf8mb4 COLLATE utf8mb4_turkish_ci`).
  - Ensure data is decoded properly before processing and re-encoded using strictly UTF-8 before transmission.

---

## 4. JSON Payload Encoding (RFC 8259 Compliance)
- All generated JSON bodies must preserve native UTF-8 strings or use Unicode escape sequences (e.g., `\u00e7` for `ç`) if the transport layer strictly requires it.
- In PHP, always serialize JSON with:
  ```php
  json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_THROW_ON_ERROR);
  ```
- Never strip or modify accents, umlauts, or Turkish diacritics under any circumstances during JSON serialization.

---

## 5. Error Mitigation & Mojibake Auto-Repair
If an incoming payload or database column contains legacy Windows-1254 / ISO-8859-9 corrupted characters (mojibake), automatically map them to their correct Turkish UTF-8 counterparts before executing further business logic:
- `Ý` → `İ`
- `Ð` → `Ğ`
- `Þ` → `Ş`
- `ý` → `ı`
- `ð` → `ğ`
- `þ` → `ş`
