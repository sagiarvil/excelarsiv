/**
 * language_synthesizer.js — Antigravity Otonom Dil Sentezleme ve Ekosistem Genişletme Motoru
 * 
 * Bu araç, eklentiye tek bir komutla yeni bir yazılım dili entegre eder:
 * 1. Dilin katı kurallarını ve best practice'lerini belirler (rules/<lang>-rules.md)
 * 2. Yüksek yoğunluklu, az token harcayan İngilizce talimatlı beceri yazar (skills/ag-<lang>-developer/SKILL.md)
 * 3. İlgili uzman ajanı oluşturur (agents/<lang>-developer/agent.md)
 * 4. Çift Dilli Mimari: Ajana talimatlar öz İngilizce (Token tasarrufu & sıfır halüsinasyon),
 *    kullanıcıya tüm çıktılar istisnasız %100 Türkçe!
 * 5. Otomatik olarak self_updater.js ile sürümü artırır ve agy CLI ile sisteme tescil eder.
 * 
 * Kullanım:
 *   node scripts/language_synthesizer.js python
 *   node scripts/language_synthesizer.js typescript
 *   node scripts/language_synthesizer.js golang
 *   node scripts/language_synthesizer.js rust
 *   node scripts/language_synthesizer.js <herhangi_bir_dil>
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const SCRIPT_DIR = __dirname;
const PLUGIN_ROOT = path.resolve(SCRIPT_DIR, '..');
const MANIFEST_PATH = path.join(PLUGIN_ROOT, 'plugin.json');

const args = process.argv.slice(2);
if (args.length === 0 || args[0].startsWith('-')) {
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('🧬 ANTIGRAVITY OTONOM DİL SENTEZLEME MOTORU (LANGUAGE SYNTHESIZER)');
  console.log('Kullanım: node scripts/language_synthesizer.js <dil_adi>');
  console.log('Örnekler: python, typescript, golang, rust, csharp, ruby, kotlin');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  process.exit(0);
}

const rawLang = args[0].toLowerCase().trim();
const langMap = {
  'py': 'python',
  'python': 'python',
  'ts': 'typescript',
  'typescript': 'typescript',
  'js': 'javascript',
  'javascript': 'javascript',
  'go': 'golang',
  'golang': 'golang',
  'rs': 'rust',
  'rust': 'rust',
  'cs': 'csharp',
  'csharp': 'csharp',
  'c#': 'csharp',
  'rb': 'ruby',
  'ruby': 'ruby',
  'kt': 'kotlin',
  'kotlin': 'kotlin',
  'swift': 'swift'
};

const lang = langMap[rawLang] || rawLang;
const langUpper = lang.charAt(0).toUpperCase() + lang.slice(1);
const skillName = `ag-${lang}-developer`;
const agentName = `${lang}-developer`;

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log(`🧬 YENİ YAZILIM DİLİ SENTEZLENİYOR: [${langUpper.toUpperCase()}]`);
console.log(`Hedef Eklenti : ${PLUGIN_ROOT}`);
console.log(`Beceri Adı    : ${skillName}`);
console.log(`Ajan Adı      : ${agentName}`);
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

// Dil profilleri ve katı invariantlar
const profiles = {
  python: {
    title: 'Python 3.12/3.14 Master Developer',
    desc: 'Python master skill. Enforces strict type annotations, PEP 8, ruff/flake8, defensive error handling, and pytest verification.',
    tooling: 'Python 3.14 / pytest / ruff / mypy',
    invariants: [
      'Strict Type Annotations: Every function must have explicit parameter types and return type annotations (def func(x: int) -> str:).',
      'Defensive None Handling: Always guard optional values with Optional[T] / T | None and explicit checks (if val is not None:).',
      'Modern Python Idioms: Prefer f-strings, dataclasses, context managers (with), and pathlib over os.path where possible.',
      'UTF-8 Clean Standard: All file I/O operations MUST explicitly state encoding="utf-8". Never leave encoding implicit.',
      'Security Invariants: Never use eval(), exec(), or insecure pickle.loads(). Always sanitize user inputs and use parameterized SQL queries.',
      'Two-Stage Verification: Verify syntax via python -m py_compile <file> and run pytest tests before reporting completion.'
    ]
  },
  typescript: {
    title: 'TypeScript & Modern Node.js Master Developer',
    desc: 'TypeScript & Node.js master skill. Enforces strict mode, zero any, explicit interfaces, ESLint compliance, and Vitest/Jest verification.',
    tooling: 'TypeScript 5+ / Node.js / Vitest / ESLint',
    invariants: [
      'Strict Mode & Zero Any: TypeScript compiler strict: true is mandatory. The `any` type is strictly forbidden; use `unknown` with type narrowing.',
      'Explicit Interfaces & Types: All function signatures, props, API payloads, and database models must have explicit types or interfaces.',
      'Defensive Null Safety: Always use optional chaining (?.) and nullish coalescing (??). Loose equality (==) is strictly forbidden; use ===.',
      'Async/Await Discipline: All asynchronous operations must be handled with async/await and wrapped in defensive try/catch blocks.',
      'Security Invariants: Prevent prototype pollution, sanitize DOM injections, avoid eval()/new Function(), and guard regex against ReDoS.',
      'Two-Stage Verification: Verify syntax and types via npx tsc --noEmit and node --check before reporting completion.'
    ]
  },
  golang: {
    title: 'Go / Golang Systems & Backend Architect',
    desc: 'Go backend architect skill. Enforces idiomatic Go, strict error propagation, concurrency safety (mutex/channels), and go vet / golangci-lint.',
    tooling: 'Go 1.22+ / gofmt / go vet / golangci-lint',
    invariants: [
      'Explicit Error Handling: Never ignore errors with `_`. Always check `if err != nil` and wrap with `fmt.Errorf("context: %w", err)`.',
      'Concurrency Safety: Protect shared state with sync.Mutex / sync.RWMutex. Avoid goroutine leaks by passing context.Context with timeout.',
      'Zero Panic in Libraries: Panic is only allowed during startup initialization. Use error returns for normal runtime error states.',
      'Idiomatic Formatting: Code must pass `gofmt -s -w` and `go vet ./...` with zero warnings.',
      'Resource Management: Always defer resource cleanup immediately after acquisition (defer resp.Body.Close(), defer file.Close()).',
      'Two-Stage Verification: Run `go vet` and `go test -v ./...` before reporting completion.'
    ]
  },
  rust: {
    title: 'Rust High-Performance & Systems Engineer',
    desc: 'Rust master skill. Enforces borrow checker discipline, zero unsafe (unless verified), Result/Option propagation, and clippy compliance.',
    tooling: 'Rust 2024 / cargo / rustfmt / clippy',
    invariants: [
      'Zero Unsafe: The `unsafe` block is strictly forbidden unless explicitly requested and formally verified.',
      'Strict Result Propagation: Use `?` operator for error propagation. Avoid `.unwrap()` and `.expect()` in production code paths.',
      'Borrow Checker Discipline: Favor borrowing (&T, &mut T) over cloning unless ownership transfer is genuinely required.',
      'Clippy Clean: Code must compile with `cargo clippy --all-targets -- -D warnings` with zero warnings.',
      'Defensive Memory: Prevent resource exhaustion, validate all external input boundaries, and use RAII guard types.',
      'Two-Stage Verification: Run `cargo check` and `cargo test` before reporting completion.'
    ]
  }
};

// Bilinmeyen dil için dinamik şablon
const profile = profiles[lang] || {
  title: `${langUpper} Software Engineer`,
  desc: `${langUpper} engineering skill. Enforces strict typing, static analysis, defensive error handling, and test-driven verification.`,
  tooling: `${langUpper} Compiler / Linter / Test Runner`,
  invariants: [
    `Strict Syntax & Type Safety: Enforce explicit typing and strict language conventions for ${langUpper}.`,
    'Defensive Error Handling: Guard against null/nil pointers, unhandled exceptions, and boundary overflows.',
    'Security & OWASP Compliance: Never trust external input, use parameterized queries, and avoid unsafe reflection/eval.',
    'Idiomatic Standards: Adhere to official community style guides, standard linters, and zero-warning build policies.',
    'Two-Stage Verification: Execute language linter/compiler check and unit tests before reporting completion.'
  ]
};

// 1. KURAL DOSYASI ÜRET (rules/<lang>-rules.md)
const rulesDir = path.join(PLUGIN_ROOT, 'rules');
if (!fs.existsSync(rulesDir)) fs.mkdirSync(rulesDir, { recursive: true });

const rulesFilePath = path.join(rulesDir, `${lang}-rules.md`);
const rulesContent = `# ${langUpper.toUpperCase()} CODING & ARCHITECTURAL INVARIANTS

> **Scope:** Universal ${langUpper} Development Standard  
> **Efficiency Mandate:** Zero Preamble · Karpathy Surgical Precision · Maximum Token Efficiency  
> **User Output Rule:** Chat responses and tool summaries must remain 100% Turkish.

---

## 🏛️ CORE INVARIANTS & STANDARDS
${profile.invariants.map((inv, idx) => `${idx + 1}. **${inv.split(':')[0]}:** ${inv.split(':').slice(1).join(':').trim()}`).join('\n')}

---

## 🛡️ TOKEN CONSERVATION PROTOCOL
- **Dense English Directives:** Internal prompts, logic rules, and skills use ultra-concise English to minimize token consumption.
- **Surgical Diffs:** Never re-write entire files. Only output modified blocks or surgical edits.
- **Verification First:** Never declare success without running compiler/linter verification.
`;

fs.writeFileSync(rulesFilePath, rulesContent, 'utf8');
console.log(`📝 1. Kural dosyası oluşturuldu: rules/${lang}-rules.md`);

// 2. BECERİ DOSYASI ÜRET (skills/ag-<lang>-developer/SKILL.md)
const skillDir = path.join(PLUGIN_ROOT, 'skills', skillName);
if (!fs.existsSync(skillDir)) fs.mkdirSync(skillDir, { recursive: true });

const skillFilePath = path.join(skillDir, 'SKILL.md');
const skillContent = `---
name: ${skillName}
description: >-
  ${profile.desc}
---

# ${profile.title.toUpperCase()}

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, \`toolAction\`, and \`toolSummary\` strictly in **100% Turkish**. Never output English to the user.

This skill provides autonomous, token-efficient, production-grade engineering for ${langUpper}.

---

## 🎯 1. CORE INVARIANTS & POLICIES
${profile.invariants.map((inv, idx) => `${idx + 1}. **${inv.split(':')[0]}:** ${inv.split(':').slice(1).join(':').trim()}`).join('\n')}

---

## ⚡ 2. SURGICAL OPERATIONAL WORKFLOW
1. **Analyze First:** Read target files and establish exact dependencies before editing.
2. **Minimal Edit:** Make small, contiguous, surgical modifications with zero unnecessary preamble.
3. **Verify:** Run syntax, linter, and test checks. Clean UTF-8 BOM bytes if applicable.
4. **Report (Turkish):** Present results clearly to the user in 100% Turkish.
`;

fs.writeFileSync(skillFilePath, skillContent, 'utf8');
console.log(`📝 2. Beceri dosyası oluşturuldu: skills/${skillName}/SKILL.md`);

// 3. AJAN DOSYASI ÜRET (agents/<lang>-developer/agent.md)
const agentDir = path.join(PLUGIN_ROOT, 'agents', agentName);
if (!fs.existsSync(agentDir)) fs.mkdirSync(agentDir, { recursive: true });

const agentFilePath = path.join(agentDir, 'agent.md');
const agentContent = `---
name: ${agentName}
description: >-
  ${profile.desc}
skills: ${skillName}
skills-path: config/skills/${skillName}/
---

# ${langUpper.toUpperCase()} DEVELOPER — ROLE PROTOCOL

> [!IMPORTANT]
> **CRITICAL USER OUTPUT MANDATE:** You MUST output all user-facing chat responses, explanations, progress reports, \`toolAction\`, and \`toolSummary\` strictly in **100% Turkish**. Never output English to the user.

- **Assigned Skill:** \`${skillName}\`
- **Primary Tooling:** ${profile.tooling}
- **Operating Standard:** Zero Preamble · Minimal Token Consumption · 100% Verified Code

## 🎯 OPERATIONAL STEPS
1. **Strict Architecture:** Implement code following \`${lang}-rules.md\` invariants.
2. **Token Conservation:** Keep reasoning ultra-dense; emit only precise, surgical changes.
3. **Verification Before Completion:** Execute language linter/compiler and confirm zero errors before finishing.
`;

fs.writeFileSync(agentFilePath, agentContent, 'utf8');
console.log(`📝 3. Ajan dosyası oluşturuldu: agents/${agentName}/agent.md`);

// 4. OTONOM GÜNCELLEME VE VERSİYON ARTIRMA (self_updater.js tetikle)
console.log('\n🔄 4. Otonom Güncelleme Motoru (self_updater.js) tetikleniyor...');
const updaterScript = path.join(SCRIPT_DIR, 'self_updater.js');
if (fs.existsSync(updaterScript)) {
  try {
    const updateOut = execSync(`node "${updaterScript}" --minor "Yeni yazılım dili ekosistemi eklendi: ${langUpper} (${agentName})"`, { stdio: ['pipe', 'pipe', 'pipe'] }).toString('utf8');
    console.log(updateOut);
  } catch (err) {
    console.error('⚠️ Güncelleme motoru hatası:', err.message);
  }
}

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log(`🎉 ${langUpper.toUpperCase()} EKOSİSTEMİ BAŞARIYLA ENTEGRE EDİLDİ!`);
console.log(`• Beceri : /${skillName}`);
console.log(`• Ajan   : ${agentName}`);
console.log(`• Kural  : rules/${lang}-rules.md`);
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
