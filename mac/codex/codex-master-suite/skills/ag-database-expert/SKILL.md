---
name: ag-database-expert
description: Database architecture and DBA skill. Covers 3NF normalization, index optimization, two-step migrations, and ACID transactions.
---

# DATABASE ARCHITECT & DBA MASTER SKILL

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, `toolAction`, and `toolSummary` strictly in **100% Turkish**. Never output English to the user.

---

## 🗄️ 1. SCHEMA & QUERY INVARIANTS
1. **Zero `SELECT *` in Production:** Always query specific columns to preserve memory and network bandwidth.
2. **Indexing Invariant:** Index all foreign keys and filter columns frequently used in `WHERE` and `JOIN` clauses.
3. **ACID Transactions:** Wrap multi-table state transitions in transactions with rollback handlers.
4. **Deterministic Migration:** Apply database changes exclusively through versioned Up/Down migration scripts.

### Türkçe Collation & Karakter Kodlaması Standardı (Zorunlu P0)
- Tüm MySQL / MariaDB tabloları ve kolonları `utf8mb4_turkish_ci` (veya `utf8_turkish_ci`) collation ile yapılandırılır.
- Veritabanı bağlantısı açılır açılmaz `SET NAMES utf8mb4 COLLATE utf8mb4_turkish_ci` sorgusu zorunludur.
- Veritabanından çekilen veya yazılan Türkçe metinlerde mojibake (`Ý, Ð, Þ, ý, ð, þ`) otomatik onarılır.
