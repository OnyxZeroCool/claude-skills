# vba-xlsm

Claude Code skill for creating Excel `.xlsm` files with VBA macros — pure Python, no Excel required.

## What it does

- Converts `.xlsx` → `.xlsm` by injecting VBA modules
- Generates `vbaProject.bin` from scratch (implements MS-CFB v3 + MS-OVBA compression in Python stdlib)
- Handles standard modules (`.bas`) and document modules (`.cls`)

## Dependencies

```bash
pip install oletools   # for validation only; not required at runtime
```

No other external dependencies — `vba_project_builder.py` uses only Python stdlib.

## Usage

```bash
# Single module
python3 vba-inject.py input.xlsx output.xlsm Module1.bas

# Multiple modules
python3 vba-inject.py input.xlsx output.xlsm Module1.bas Module2.bas ThisWorkbook.cls
```

### As a library

```python
from vba_project_builder import build_vba_project

modules = [
    ("ThisWorkbook", "document", ""),
    ("Module1", "standard", 'Sub Hello()\n    MsgBox "Hello!"\nEnd Sub'),
]
build_vba_project(modules, "vbaProject.bin", code_page=1251)
```

## Files

| File | Description |
|------|-------------|
| `vba-inject.py` | CLI tool: xlsx + .bas/.cls → xlsm |
| `vba_project_builder.py` | Core library: builds `vbaProject.bin` |
| `SKILL.md` | Claude Code skill definition |
| `xlsm-reference.md` | Technical reference (MS-OVBA / MS-CFB internals) |

## Install as Claude Code skill

```bash
cp -r vba-xlsm ~/.claude/skills/
```

Then add to `~/.claude/settings.json`:
```json
{
  "skills": ["~/.claude/skills/vba-xlsm/SKILL.md"]
}
```
