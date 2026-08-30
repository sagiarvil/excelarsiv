# Production Functions Recovery

This file intentionally lives under `functions/` so the production runtime detector performs one full backend redeploy after the 2026-08-30 Hosting release exposed missing Gen2 endpoints.

The permanent fix is in `.github/workflows/deploy-firebase.yml`: a runtime release now redeploys the Functions set when required production endpoints are missing, even when backend source code itself is unchanged.
