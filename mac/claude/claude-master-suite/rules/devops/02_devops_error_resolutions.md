# DEVOPS INCIDENT RESPONSE & TROUBLESHOOTING PLAYBOOKS

This guide details remediation protocols for secret exposure, container crashes, and permission failures.

## 1. SECRET LEAK RESPONSE PROTOCOL
1. Immediately revoke and rotate the exposed API key or credential in the upstream provider console.
2. Add the leaked path pattern to `.gitignore`.
3. Purge history using BFG Repo-Cleaner or git filter-repo before pushing.

## 2. CONTAINER CRASH (CRASHLOOPBACKOFF) DIAGNOSTICS
1. Inspect termination logs: `docker logs --tail 100 <container_name>`.
2. Verify port bindings and environment variable completeness.
3. Validate internal file ownership: `chown -R www-data:www-data /var/www/html`.
