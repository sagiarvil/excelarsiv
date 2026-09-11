# PYTHON ERROR RESOLUTIONS & RESOURCE PATTERNS

This guide details deterministic solutions for memory leaks, mutable argument pitfalls, and async event loop blocking.

## 1. MUTABLE DEFAULT ARGUMENT PATTERN
```python
# CORRECT IMPLEMENTATION:
from typing import Optional

def add_task(task: str, task_list: Optional[list[str]] = None) -> list[str]:
    if task_list is None:
        task_list = []
    task_list.append(task)
    return task_list
```

## 2. CROSS-PLATFORM PATHLIB MANAGEMENT
```python
from pathlib import Path

target_file = Path("storage") / "reports" / "summary.json"
target_file.parent.mkdir(parents=True, exist_ok=True)
with target_file.open("w", encoding="utf-8") as f:
    json.dump(payload, f, ensure_ascii=False, indent=2)
```

## 3. NON-BLOCKING ASYNC HTTP
```python
import httpx

async def fetch_api(url: str) -> dict:
    async with httpx.AsyncClient(timeout=10.0) as client:
        resp = await client.get(url)
        resp.raise_for_status()
        return resp.json()
```
