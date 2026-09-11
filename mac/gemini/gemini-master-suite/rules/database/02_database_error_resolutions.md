# DATABASE OPTIMIZATION & DEADLOCK RESOLUTION PLAYBOOKS

This guide details diagnostics for query bottlenecks, Full Table Scans, and transaction lock contention.

## 1. SLOW QUERY & FULL TABLE SCAN RESOLUTION
- Prepend `EXPLAIN ANALYZE` to identify Sequential Scans (`Seq Scan` / `ALL`).
- Construct targeted composite indexes:
  ```sql
  CREATE INDEX idx_orders_user_status ON orders (user_id, status, created_at DESC);
  ```

## 2. DEADLOCK PREVENTION STRATEGY
- Always acquire locks and access tables in a deterministic, identical global order across all application services (e.g., always update `accounts` before `transactions`).
- Keep transaction lifespans minimal; NEVER make external HTTP calls inside a database transaction.
