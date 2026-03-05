# summarize

Claude Code skill for summarizing local files and web pages with follow-up Q&A.

## What it does

- Reads local files (`.txt`, `.md`, `.html`, `.pdf`) or fetches web pages
- Generates structured summaries at three detail levels
- Supports multilingual output
- Enters interactive Q&A mode after the summary

## Dependencies

No external dependencies — uses Claude Code's built-in `Read` and `WebFetch` tools.

## Usage

```
/summarize <path_or_url> [mode] [lang=XX]
```

| Argument | Description |
|----------|-------------|
| `path_or_url` | Local file path or `https://` URL |
| `mode` | `short`, `normal` (default), `detailed` |
| `lang=XX` | Output language code (e.g. `lang=en`, `lang=de`). Defaults to source language |

### Examples

```
/summarize report.pdf
/summarize paper.pdf detailed lang=en
/summarize https://example.com/article short
```

## Install as Claude Code skill

```bash
cp -r summarize ~/.claude/skills/
```

Add to `~/.claude/settings.json`:
```json
{
  "skills": ["~/.claude/skills/summarize/SKILL.md"]
}
```

## Files

| File | Description |
|------|-------------|
| `SKILL.md` | Claude Code skill definition |
