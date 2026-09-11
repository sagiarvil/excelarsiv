---
name: ag-php-developer
description: Full-Stack PHP & Web Developer Master Skill. Enforces modern PHP 8.3/8.4, strict typing, OWASP security, and clean architecture.
---

# PHP & WEB DEVELOPER MASTER SKILL

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, `toolAction`, and `toolSummary` strictly in **100% Turkish**. Never output English to the user.

This master skill unifies backend architecture, database resilience, frontend layouts, and PHP security under a single deterministic standard.

---

## 🐘 1. CORE PHP 8.3/8.4 INVARIANTS
1. **Strict Types on Line 1:** Every PHP file MUST begin with `<?php` followed immediately by `declare(strict_types=1);`.
2. **Explicit Type Signatures:** All function parameters, return values, and class properties must be typed.
3. **Implicit Nullable Ban (PHP 8.4):** Signatures like `string $x = null` are forbidden. Use `?string $x = null`.
4. **Strict Comparisons:** Always use `===` and `!==`. Loose equality is strictly forbidden.
5. **Defensive Null Safety:** Access arrays via `$arr['key'] ?? $default` and objects via nullsafe `$obj?->prop?->method()`.

---

## 🔒 2. SECURITY & OWASP INVARIANTS
1. **100% Prepared Statements:** Raw SQL string concatenation is forbidden. Always use PDO prepared statements.
2. **XSS Protection:** Escape all dynamic values rendered into HTML using `htmlspecialchars($v, ENT_QUOTES, 'UTF-8')`.
3. **Mass Assignment Prevention:** Never pass raw request data (`$request->all()`) directly to ORM model creation methods.
4. **Banned Functions:** The `@` error suppression operator, `eval()`, `extract()`, and `$$var` are strictly banned.

---

## ⚡ 3. PERFORMANCE & TWO-STAGE VERIFICATION
1. **Zero N+1 Queries:** Database queries inside loops are forbidden. Always eager load via `with()`.
2. **Two-Stage Verification Hook:**
   - Step 1: Run `node .../bom_utf8_scan.js --fix <file>` to strip BOM bytes.
   - Step 2: Run `php -l <file>` to verify syntax. Never mark tasks complete without physical proof.
