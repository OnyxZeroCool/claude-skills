# transcribe

Claude Code skill for transcribing audio files and generating structured DOCX summaries from templates.

## What it does

- Transcribes `.mp3`, `.wav`, `.m4a` using mlx-whisper (Apple Silicon GPU)
- Corrects transcript via Claude: fixes misrecognized words, punctuation, sentence boundaries
- Optionally fills a DOCX template with structured content extracted from the transcript
- Style guidance from reference `.docx` documents (`--ref` flag)

## Dependencies

```bash
pip install mlx-whisper python-docx
```

> **Note:** `mlx-whisper` requires Apple Silicon (M1/M2/M3). On Intel/Linux, replace with `faster-whisper` and update the `transcribe()` function in `process_docx.py`.

## Setup

1. Copy skill to `~/.claude/skills/`:
   ```bash
   cp -r transcribe ~/.claude/skills/
   ```

2. Add to `~/.claude/settings.json`:
   ```json
   {
     "skills": ["~/.claude/skills/transcribe/SKILL.md"]
   }
   ```

3. *(Optional)* Add DOCX templates:
   ```bash
   mkdir ~/.claude/skills/transcribe/templates
   cp my_template.docx ~/.claude/skills/transcribe/templates/retro.docx
   ```

4. *(Optional)* Add style references:
   ```bash
   mkdir ~/.claude/skills/transcribe/references
   cp style_example.docx ~/.claude/skills/transcribe/references/
   ```

## Usage

```
/transcribe <path_to_mp3>                        → corrected transcript (.txt)
/transcribe <path_to_mp3> --ref                   → transcript with style guidance (.txt)
/transcribe <path_to_mp3> <template_name>         → DOCX from template
/transcribe <path_to_mp3> <template_name> --ref   → DOCX from template + style references
```

### Examples

```
/transcribe ~/meetings/standup.mp3
/transcribe ~/interviews/qa_session.m4a retro --ref
```

## Files

| File | Description |
|------|-------------|
| `SKILL.md` | Claude Code skill definition |
| `process_docx.py` | CLI tool: transcribe audio, extract/fill DOCX templates, read references |

## DOCX Template Format

Templates are standard `.docx` files with placeholder text like `[Date]`, `[Participants]`, etc. Claude extracts the template structure, then fills it with content from the transcript. See `SKILL.md` for the JSON content format.
