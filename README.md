# Claude Code Skills

A collection of skills for [Claude Code](https://claude.ai/claude-code) — Anthropic's CLI coding assistant.

## Installation

Copy the desired skill folder into `~/.claude/skills/`:

```bash
cp -r vba-xlsm ~/.claude/skills/
```

Then reference it in your `~/.claude/settings.json`:

```json
{
  "skills": ["~/.claude/skills/vba-xlsm/SKILL.md"]
}
```

## Skills

### [vba-xlsm](./vba-xlsm/)

Create Excel `.xlsm` files with VBA macros programmatically — pure Python, no Excel required.

- Inject VBA modules into existing `.xlsx` / `.xlsm` files
- Generate `vbaProject.bin` from scratch (MS-CFB + MS-OVBA implemented in stdlib)
- No external dependencies except `oletools` for validation

```bash
python3 vba-inject.py input.xlsx output.xlsm Module1.bas
```

---

### [youtube](./youtube/)

Analyze YouTube videos: transcript, summary, comments, and Q&A.

- Fetches subtitles (manual or auto-generated), English by default
- Generates structured summary via Claude
- Analyzes top comments by likes
- Enters Q&A mode after summary

```
/youtube https://youtu.be/... [summary|comments|full]
```

**Requirements:** `yt-dlp`, `jq`

---

### [summarize](./summarize/)

Summarize local files and web pages with follow-up Q&A.

- Supports `.txt`, `.md`, `.html`, `.pdf` and URLs
- Three detail levels: `short`, `normal`, `detailed`
- Multilingual output via `lang=XX` parameter

```
/summarize report.pdf detailed lang=en
/summarize https://example.com/article
```

---

### [transcribe](./transcribe/)

Transcribe audio files and generate structured DOCX summaries from templates.

- Transcribes `.mp3`, `.wav`, `.m4a` using mlx-whisper (Apple Silicon)
- Corrects transcript via Claude (fixes misrecognized words, punctuation)
- Fills DOCX templates with structured content from the transcript
- Style guidance from reference documents

```
/transcribe meeting.mp3
/transcribe interview.mp3 retro
```

**Requirements:** `mlx-whisper` (Apple Silicon), `python-docx`
