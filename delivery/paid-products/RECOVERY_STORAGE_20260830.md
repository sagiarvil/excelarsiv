# Paid Storage Recovery — 2026-08-30

This marker intentionally triggers the private paid-product delivery layer once after production checkout correctly returned `PRODUCT_NOT_READY` for a catalog item whose source binary exists in the repository but was missing from Firebase Storage.

The production workflow is hardened in the same change so every future runtime release runs the idempotent paid-product sync/parity gate and self-heals missing or mismatched private delivery objects before accepting checkout traffic.
