# PYTHON CODING & ARCHITECTURAL INVARIANTS

> **Scope:** Universal Python Development Standard  
> **Efficiency Mandate:** Zero Preamble · Karpathy Surgical Precision · Maximum Token Efficiency  
> **User Output Rule:** Chat responses and tool summaries must remain 100% Turkish.

---

## 🏛️ CORE INVARIANTS & STANDARDS
1. **Strict Type Annotations:** Every function must have explicit parameter types and return type annotations (def func(x: int) -> str:).
2. **Defensive None Handling:** Always guard optional values with Optional[T] / T | None and explicit checks (if val is not None:).
3. **Modern Python Idioms:** Prefer f-strings, dataclasses, context managers (with), and pathlib over os.path where possible.
4. **UTF-8 Clean Standard:** All file I/O operations MUST explicitly state encoding="utf-8". Never leave encoding implicit.
5. **Security Invariants:** Never use eval(), exec(), or insecure pickle.loads(). Always sanitize user inputs and use parameterized SQL queries.
6. **Two-Stage Verification:** Verify syntax via python -m py_compile <file> and run pytest tests before reporting completion.

---

## 🛡️ TOKEN CONSERVATION PROTOCOL
- **Dense English Directives:** Internal prompts, logic rules, and skills use ultra-concise English to minimize token consumption.
- **Surgical Diffs:** Never re-write entire files. Only output modified blocks or surgical edits.
- **Verification First:** Never declare success without running compiler/linter verification.
