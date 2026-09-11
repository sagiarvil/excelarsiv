# PHP OWASP SECURITY INVARIANTS & HARDENING RULES

This standard defines mandatory security protocols to eliminate OWASP Top 10 vulnerabilities in PHP codebases.

## 1. SQL INJECTION ELIMINATION (100% PREPARED STATEMENTS)
1. **String Concatenation Ban:** Constructing SQL queries via string concatenation, variable interpolation (`"WHERE id = " . $id`), or raw format strings is STRICTLY FORBIDDEN.
2. **PDO Prepared Statements:** All database queries MUST use parameterized PDO prepared statements:
   ```php
   // CORRECT:
   $stmt = $pdo->prepare("SELECT id, name, email FROM users WHERE email = :email AND status = :status");
   $stmt->execute([':email' => $email, ':status' => 'active']);
   ```

## 2. XSS (CROSS-SITE SCRIPTING) PREVENTION
1. **Unescaped Echo Ban:** Directly outputting unescaped variables into HTML context (`echo $input;`) is FORBIDDEN.
2. **Mandatory Escaping:** All dynamic strings rendered in HTML must pass through `htmlspecialchars()` with explicit parameters:
   ```php
   echo htmlspecialchars($content ?? '', ENT_QUOTES, 'UTF-8');
   ```

## 3. MASS ASSIGNMENT PROTECTION
- When using ORMs (e.g., Eloquent), passing raw request arrays (`$request->all()`) directly into `Model::create()` or `update()` is FORBIDDEN.
- Always use validated datasets (`$request->validated()`) or explicit attribute whitelists (`$request->only([...])`).

## 4. FORBIDDEN DANGEROUS FUNCTIONS & PATTERNS
- **Error Suppression Operator (`@`):** BANNED. Never suppress errors with `@`. Handle errors defensively or via structured `try/catch`.
- **Arbitrary Code Execution:** `eval()`, `create_function()` are STRICTLY FORBIDDEN.
- **Variable Injection:** `extract()` and variable variables (`$$var`) are BANNED.
- **Insecure Deserialization:** `unserialize()` on untrusted inputs is FORBIDDEN; use `json_decode()` / `json_encode()`.
