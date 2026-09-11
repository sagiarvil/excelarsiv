# TYPESCRIPT CODING & ARCHITECTURAL INVARIANTS

> **Scope:** Universal Typescript Development Standard  
> **Efficiency Mandate:** Zero Preamble · Karpathy Surgical Precision · Maximum Token Efficiency  
> **User Output Rule:** Chat responses and tool summaries must remain 100% Turkish.

---

## 🏛️ CORE INVARIANTS & STANDARDS
1. **Strict Mode & Zero Any:** TypeScript compiler strict: true is mandatory. The `any` type is strictly forbidden; use `unknown` with type narrowing.
2. **Explicit Interfaces & Types:** All function signatures, props, API payloads, and database models must have explicit types or interfaces.
3. **Defensive Null Safety:** Always use optional chaining (?.) and nullish coalescing (??). Loose equality (==) is strictly forbidden; use ===.
4. **Async/Await Discipline:** All asynchronous operations must be handled with async/await and wrapped in defensive try/catch blocks.
5. **Security Invariants:** Prevent prototype pollution, sanitize DOM injections, avoid eval()/new Function(), and guard regex against ReDoS.
6. **Two-Stage Verification:** Verify syntax and types via npx tsc --noEmit and node --check before reporting completion.

---

## 🛡️ TOKEN CONSERVATION PROTOCOL
- **Dense English Directives:** Internal prompts, logic rules, and skills use ultra-concise English to minimize token consumption.
- **Surgical Diffs:** Never re-write entire files. Only output modified blocks or surgical edits.
- **Verification First:** Never declare success without running compiler/linter verification.
