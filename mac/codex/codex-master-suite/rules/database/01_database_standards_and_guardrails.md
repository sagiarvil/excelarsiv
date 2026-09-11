# DATABASE ARCHITECTURE & SQL INTEGRITY STANDARDS

This rulebook enforces relational integrity, indexing strategies, and database safety invariants.

## 1. STRICTLY FORBIDDEN DATABASE PATTERNS
1. **`SELECT *` Ban in Production:** Querying with `SELECT *` on production tables is FORBIDDEN. Explicitly declare only the required columns to save memory and I/O.
2. **Unindexed Foreign Keys & Filter Columns:** Foreign keys and columns frequently used in `WHERE`, `JOIN`, and `ORDER BY` must not remain unindexed.
3. **Manual Production Schema Mutation:** Applying manual database modifications via GUI or CLI without two-step (Up/Down) versioned migration scripts is STRICTLY FORBIDDEN.
4. **Multi-Table Operations Without Transactions:** Mutating multiple related tables (e.g., balance transfers, inventory deductions) without ACID transactions is BANNED.
5. **Timezone-Agnostic Timestamps:** Storing datetimes without timezone data is FORBIDDEN. Use UTC standard and `timestamptz` or `DATETIME` with UTC normalization.
6. **Plaintext Password Storage:** Storing credentials or sensitive tokens in plaintext is a catastrophic security violation. Always use Argon2id or bcrypt with salt.
