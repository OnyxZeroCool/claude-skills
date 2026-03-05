# youtube

Claude Code skill for analyzing YouTube videos — transcript, summary, comments, and Q&A.

## What it does

- Fetches video subtitles (manual > auto-generated, English by default)
- Generates a structured summary using Claude
- Analyzes top comments sorted by likes
- Enters interactive Q&A mode after the summary

## Dependencies

```bash
brew install yt-dlp jq
```

## Usage

```
/youtube <url> [mode]
```

| Mode | Description |
|------|-------------|
| `summary` (default) | Transcript + structured summary + Q&A |
| `comments` | Top comment analysis + Q&A |
| `full` | Transcript + comments combined + Q&A |

## Install as Claude Code skill

```bash
cp -r youtube ~/.claude/skills/
```

Add to `~/.claude/settings.json`:
```json
{
  "skills": ["~/.claude/skills/youtube/SKILL.md"]
}
```

## Files

| File | Description |
|------|-------------|
| `SKILL.md` | Claude Code skill definition |
