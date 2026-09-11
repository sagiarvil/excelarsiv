# PYTHON 3.12+ CODING STANDARDS & NEGATIVE CONSTRAINTS

This standard outlines strict rules for Python scripting, data engineering, automation, and AI integrations.

## 1. STRICTLY FORBIDDEN PYTHON PATTERNS
1. **Bare Except Ban:** Writing `except:` without an explicit exception class is STRICTLY FORBIDDEN. It catches `SystemExit` and `KeyboardInterrupt`, masking fatal bugs. Catch specific exceptions or log `except Exception as e:`.
2. **Mutable Default Argument Ban:** Using empty lists or dictionaries as default arguments (`def fn(items=[]):`) is FORBIDDEN. Use `None` as default and initialize internally.
3. **Unmanaged Resource Access:** Calling `open()` without a context manager (`with open(...) as f:`) is FORBIDDEN.
4. **Missing Type Hints:** Function parameters and return values must specify explicit type annotations.
5. **Wildcard Imports:** `from module import *` is FORBIDDEN to prevent namespace pollution.
6. **Hardcoded Platform Paths:** Platform-specific path strings (`"C:\\dir\\file.txt"`) are BANNED. Always use `pathlib.Path`.
7. **Blocking I/O in Asyncio:** Synchronous calls (`time.sleep`, synchronous `requests.get`) inside `async def` are FORBIDDEN. Use `asyncio.sleep` and `httpx.AsyncClient`.
