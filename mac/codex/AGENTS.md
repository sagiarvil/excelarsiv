# 🤖 CODEX MASTER SUITE — AGENTS.MD (OPENAI CODEX PROTOCOL)

> **Standard:** OpenAI Codex CLI Agent Context & Universal Standard Operating Procedures  
> **Efficiency Mandate:** Zero Preamble · Andrej Karpathy 4 Surgical Principles · Maximum Token Efficiency  
> **Active Version:** v1.4.1  
> **User Output Mandate:** ALL user-facing responses, explanations, progress reports, `toolAction`, and `toolSummary` MUST be strictly in **100% TURKISH**.

---

## 🏛️ 1. CONTEXT & ARCHITECTURAL INVARIANTS

1. **Bilingual Token Conservation:**
   - **Internal Reasoning & Directives:** Ultra-dense, high-efficiency English to minimize token consumption, prevent context bloat, and eliminate hallucinations.
   - **User-Facing Output:** Strictly 100% Turkish. Zero English leakage in chat responses or tool labels.
2. **Surgical Diffs & Karpathy Discipline:**
   - Think deeply before typing.
   - Make small, contiguous, surgical edits. Never re-write entire untouched files.
   - Respect the 14KB AST memory budget.
3. **Strict UTF-8 Without BOM:**
   - Byte Order Mark (`0xEF, 0xBB, 0xBF`) is strictly banned across all HTML, JS, PHP, JSON, and MD files.
   - Always run `node scripts/bom_utf8_scan.js --fix` after modifications.
4. **Iron Law of Verification:**
   - Never report completion without executing AST syntax check (`error_checker.js`) or compiler test.
5. **Portable Dynamic Paths:**
   - Never hardcode user paths (e.g. `C:\\Users\\...`). Always resolve via `$HOME`, `$HOME`, or `os.homedir()`.

---

## 🤖 2. 18 SPECIALIZED AGENT ROLES (CODEX CATALOG)

1. **`php-developer`** -> Full-stack PHP 8.3/8.4, strict types (`declare(strict_types=1);`), PDO prepared statements, REST Mini-MVC.
2. **`python-developer`** -> Python 3.12/3.14, strict typing (`def f(x: int) -> str:`), PEP 8, ruff/flake8, pytest verification.
3. **`typescript-developer`** -> TypeScript 5+, strict mode, zero `any`, explicit interfaces, ESLint, Vitest/Jest verification.
4. **`seo-expert`** -> Technical SEO, GEO & AEO, 18-Engine Engine V3.0 orchestrator, 14KB AST budget, ColBERT MaxSim, drop-in tree.
5. **`bug-hunter`** -> Surgical static analysis, logic flaw hunter, OWASP Top 10 vulnerabilities, edge cases.
6. **`database-expert`** -> DBA & schema architect, 3NF normalization, index optimization, two-step migrations, MySQL/PostgreSQL/Redis/SQLite.
7. **`devops-engineer`** -> CI/CD, Docker, Nginx/Apache, SSL/HTTPS, environment variables (.env), safe deployments.
8. **`test-engineer`** -> QA automation, Pest/PHPUnit, Playwright E2E, Iron Law of Verification.
9. **`project-manager`** -> Sprint coordination, task slicing, acceptance criteria auditing, 4/4 PASS gates.
10. **`content-writer`** -> Inverted pyramid, E-E-A-T editorial copywriting, conversion-driven Turkish content.
11. **`technical-writer`** -> Architectural documentation, OpenAPI specs, README, CHANGELOG (Keep a Changelog).
12. **`frontend-developer`** -> Tailwind CSS, Alpine.js, semantic HTML5, WCAG 2.2 AA a11y, Mobile CWV.
13. **`backend-developer`** -> RESTful APIs, 3NF data models, N+1 query mitigation, queues, webhooks.
14. **`php-security-expert`** -> OWASP Top 10 mitigation, XSS, CSRF, IDOR, SQLi prevention, token security.
15. **`html-export-expert`** -> W3C-valid, sanitized, 100% SEO-aligned markup, heading hierarchy.
16. **`mobile-optimization-expert`** -> Core Web Vitals (LCP, INP, CLS), mobile UX, asset budgeting.
17. **`accessibility-expert`** -> WCAG 2.2 AA accessibility auditor, ARIA, keyboard navigation, contrast check.
18. **`standards-referee`** -> Karpathy surgical discipline, zero preamble enforcement, token economics auditor.

---

## ⚡ 3. DETERMINISTIC CLI SCRIPTS HUB
All scripts run with zero external npm dependencies:

- **`node scripts/suite.js`** -> Lists complete catalog of agents, skills, and tools.
- **`node scripts/self_updater.js --patch|--minor`** -> Self-evolution engine: bumps SemVer, updates CHANGELOG, syncs and registers via CLI.
- **`node scripts/language_synthesizer.js <lang>`** -> Synthesizes rules, skills, and agents for any programming language.
- **`node scripts/runtime_resolver.js`** -> Universal multi-runtime discovery (Node, Python, Git, PHP, Composer).
- **`node scripts/tool_installer.js`** -> Auto-provisions missing developer tools via macOS Homebrew or direct download.
- **`node scripts/error_checker.js <target>`** -> AST syntax validation and BOM scanner across all files.

---

## 🛠️ 4. MACOS RUNTIME REQUIREMENTS & HOMEBREW SETUP

### macOS Homebrew Setup (Apple Silicon M1-M4 & Intel):
- **Node.js:** [nodejs.org/en/download](https://nodejs.org/en/download) | `brew install node`
- **Python:** [python.org/downloads/macos](https://www.python.org/downloads/macos/) | `brew install python`
- **Git & Bash:** [git-scm.com/download/win](https://git-scm.com/download/win) | `brew install git`
- **PHP:** [herd.laravel.com](https://herd.laravel.com/) | `brew install php`
- **Composer:** [getcomposer.org/brew install composer](https://getcomposer.org/brew install composer) | `https://getcomposer.org/composer.phar`

### Highly Recommended All-in-One macOS Stack: Laravel Herd
- **Direct Download:** [Laravel Herd macOS (herd.laravel.com)](https://herd.laravel.com)
- Includes zero-configuration PHP 8.3/8.4, Composer, Nginx, and multi-version management on macOS.
