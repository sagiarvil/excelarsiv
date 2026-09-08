# Production Functions Recovery

This file intentionally lives under `functions/` so the production runtime detector performs one full backend redeploy after the 2026-08-30 Hosting release exposed missing Gen2 endpoints.

The permanent fix is in `.github/workflows/deploy-firebase.yml`: a runtime release now redeploys the Functions set when required production endpoints are missing, even when backend source code itself is unchanged.

## Live Production Verification (2026-09-08)
- All 8 Gen2 commerce and delivery endpoints deployed to `europe-west1`:
  `createCheckout`, `checkoutStatus`, `verifyShopierOrder`, `recoverPurchase`, `createDownloadToken`, `downloadFile`, `requestProofDemo`, `downloadProofDemo`.
- Hosting rewrites connected with zero missing endpoint warnings.
- Live checkout API smoke-tested on `https://excelarsiv.com/api/checkout` (HTTP 201 Shopier session created).
- Full production deploy completed and operational.
