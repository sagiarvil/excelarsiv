# PHP VERIFICATION HOOKS & DETERMINISTIC RESOLUTION PLAYBOOKS

This guide defines verification workflows and deterministic root-cause solutions for common PHP runtime failures.

## 1. TWO-STAGE VERIFICATION HOOK
1. **Stage 1 (Automated BOM Sanitization):** After modifying or writing any PHP file, execute the BOM audit tool:
   `node $HOME/.gemini\config\plugins\gemini-master-suite\scripts/bom_utf8_scan.js --fix <filePath>`
2. **Stage 2 (CLI Syntax Check):** Run the canonical PHP 8.3 CLI binary to verify AST validity:
   `php -l <filePath>`
   Verification fails if the output does not contain `No syntax errors detected`.

## 2. PERFORMANCE INVARIANT: ZERO N+1 QUERIES
1. **Loop Query Ban:** Executing database queries or lazy-loading ORM relations inside `foreach`, `for`, or `while` loops is STRICTLY FORBIDDEN.
2. **Eager Loading:** Always batch fetch or eager load related entities prior to iteration:
   ```php
   // CORRECT:
   $posts = Post::with(['author', 'comments.user'])->get();
   ```
3. **Database Transactions:** Multi-table writes (orders, balances, stock updates) MUST be wrapped in ACID transactions with full rollback handling:
   ```php
   $pdo->beginTransaction();
   try {
       // execute operations
       $pdo->commit();
   } catch (\Throwable $e) {
       $pdo->rollBack();
       throw $e;
   }
   ```

## 3. ERROR RESOLUTION PLAYBOOKS
- **"Undefined array key":** Replace direct access with null coalescing: `$val = $data['key'] ?? null;`.
- **"Attempt to read property on null":** Chain the nullsafe operator: `$val = $obj?->nested?->target;`.
- **"Cannot modify header information - headers already sent":**
  1. Remove BOM bytes using `bom_utf8_scan.js --fix`.
  2. Ensure zero leading whitespace before the opening `<?php` tag.
  3. Remove the trailing `?>` close tag to prevent accidental newline leaks.
