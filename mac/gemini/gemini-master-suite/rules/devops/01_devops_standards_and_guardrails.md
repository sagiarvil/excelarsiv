# DEVOPS, INFRASTRUCTURE & DEPLOYMENT INVARIANTS

This standard enforces container security, secret management, logging, and deployment verification.

## 1. STRICTLY FORBIDDEN INFRASTRUCTURE PATTERNS
1. **Hardcoded Secrets Ban:** Embedding API keys, database credentials, or private certificates into source code, Dockerfiles, or scripts is STRICTLY FORBIDDEN. Use `.env` variables.
2. **Root User Execution Ban:** Running web services or container entrypoints as `root` is BANNED. Define dedicated low-privilege users (`www-data`, `appuser`).
3. **Unbounded Log Growth:** Writing application logs without size caps or rotation (`logrotate`) is FORBIDDEN.
4. **Missing Healthchecks:** Deploying services without automated `healthcheck` endpoints and restart policies (`restart: unless-stopped`) is BANNED.
5. **Unverified Deployments:** Deploying any codebase that has not passed automated verification gates (4/4 PASS) is strictly forbidden.
