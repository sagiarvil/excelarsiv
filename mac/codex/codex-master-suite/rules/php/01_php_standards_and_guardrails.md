# MODERN PHP 8.3/8.4 STANDARDS & DEFENSIVE GUARDRAILS

This rulebook governs all PHP development, enforcing strict type contracts, null-safety, and anti-pattern bans.

## 1. STRICT TYPING & DECLARATION
1. **Line 1 Declaration:** Every PHP file MUST begin with `<?php` on line 1, immediately followed by `declare(strict_types=1);`. No whitespace, HTML, or comment may precede or separate these statements.
2. **Explicit Type Signatures:** All function arguments, class properties, and method return values MUST be explicitly typed.
3. **Implicit Nullable Ban (PHP 8.4 Invariant):** Parameter signatures such as `function test(string $x = null)` are DEPRECATED in PHP 8.4 and STRICTLY FORBIDDEN. Use explicit nullable types: `function test(?string $x = null): void`.
4. **Strict Equality:** Loose comparisons (`==`, `!=`) are FORBIDDEN. Always utilize strict equality operators (`===`, `!==`).
5. **Terminal Return Types:** Methods returning nothing must be typed `: void`. Methods terminating execution or throwing exceptions must be typed `: never`.

## 2. DEFENSIVE NULL SAFETY & ARRAY ACCESS
1. **Array Coalescing Mandatory:** Direct unchecked array access (`$data['key']`) is FORBIDDEN due to `Undefined array key` fatal warnings. Always provide fallbacks via null coalescing: `$data['key'] ?? $default`.
2. **Nullsafe Object Chaining:** Never chain methods on nullable instances directly (`$order->getShipping()->getCountry()`). Enforce the nullsafe operator: `$order?->getShipping()?->getCountry() ?? 'TR'`.

## 3. FORMATTING & ANTI-MINIFY DISCIPLINE
1. **Anti-Minification:** Never compress, minify, or collapse PHP files into single lines.
2. **PSR-12 Standard:** Maintain strict 4-space indentation, clear control structure spacing, and multi-line readable signatures.

## 4. READ-BEFORE-WRITE PROTOCOL
- Before editing or authoring any PHP class, the agent MUST inspect the existing Model, Interface, Migration, and Service contracts. Never hallucinate nonexistent database columns, method names, or class namespaces.
