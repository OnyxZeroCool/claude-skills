---
description: "Transcribe audio and generate summary DOCX from template"
user_invocable: true
argument: "<path_to_mp3> [template_name] [--ref]"
---

# Transcribe Audio

Transcribe an audio file using mlx-whisper. Optionally generate a structured DOCX summary based on a template.

## Usage

```
/transcribe <path_to_mp3>                        → corrected transcript only (txt)
/transcribe <path_to_mp3> --ref                   → corrected transcript styled per references (txt)
/transcribe <path_to_mp3> <template_name>         → DOCX from template (no references)
/transcribe <path_to_mp3> <template_name> --ref   → DOCX from template + style references
```

- `path_to_mp3` — path to audio file (mp3, wav, m4a)
- `template_name` *(optional)* — name of template without extension (e.g. `retro`). If omitted → **txt mode**
- `--ref` *(optional)* — include style references from the reference directory. Works in both txt and docx modes

## Modes

| Arguments | Mode | Output |
|-----------|------|--------|
| `<mp3>` | txt | `<dir>/<name>.txt` — corrected transcript |
| `<mp3> --ref` | txt+ref | `<dir>/<name>.txt` — corrected transcript, terminology and style guided by references |
| `<mp3> <template>` | docx | `<dir>/<name>.docx` — summary from template |
| `<mp3> <template> --ref` | docx+ref | `<dir>/<name>.docx` — summary from template, styled per references |

## Paths

Configure once after installing the skill:

- `SCRIPT` = `~/.claude/skills/transcribe/process_docx.py` — bundled, no change needed
- `TEMPLATE_DIR` = `~/.claude/skills/transcribe/templates/` — put your `.docx` templates here
- `REFERENCES_DIR` = `~/.claude/skills/transcribe/references/` — put reference `.docx` files here (optional)

## Security

- All file paths in bash commands MUST be wrapped in single quotes to prevent shell injection (e.g. `'$path'` not `$path`)
- Never interpolate user-provided paths without quoting

---

## Algorithm: TXT mode (no template)

When only `<path_to_mp3>` is provided (with or without `--ref`).

### 1. Validate

- Verify the audio file exists at the given path
- Determine output path: `<directory_of_mp3>/<mp3_name_without_ext>.txt`

### 2. Transcribe

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --transcribe '<mp3_path>' --output_dir /tmp
```

This saves `<mp3_name>.txt` in `/tmp`. Read the generated file.

### 3. Read style references (only if `--ref` flag is present)

**Skip this step if `--ref` was NOT provided.**

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --references '~/.claude/skills/transcribe/references/'
```

This outputs the text content of all .docx files in the reference directory. Pass this to the sonnet agent in the next step as additional context for correcting terminology, names, and style.

### 4. Correct transcript (via sonnet agent)

Spawn an Agent with `model: "sonnet"` to correct the raw transcript.

Agent prompt — pass the full transcript text and these instructions:

> You are a transcript editor. Fix speech recognition errors in the following transcript:
> - Misrecognized words and homophones
> - Names of people, products, and technical terms
> - Missing or incorrect punctuation
> - Sentence boundaries and paragraph breaks
> - Remove filler words (uh, um, э, ну) unless they carry meaning
>
> Rules:
> - Fix ONLY recognition errors — do NOT change meaning, rephrase, or summarize
> - Preserve the original speaker's wording as closely as possible
> - If unsure whether something is an error, leave it as-is
> - Return ONLY the corrected transcript text, nothing else

If `--ref` was provided, add to the agent prompt:

> Use the following reference documents to guide your corrections — match terminology,
> proper names, product names, and domain-specific vocabulary from these references:
>
> <paste reference text here>

The agent returns the corrected text.

### 5. Save output

Write the corrected transcript to `<directory_of_mp3>/<mp3_name_without_ext>.txt` using the Write tool.

### 6. Cleanup

Delete the temporary file `/tmp/<mp3_name>.txt`.

### 7. Report

Tell the user the output file path.

---

## Algorithm: DOCX mode (with template)

When `<template_name>` is provided (with or without `--ref`).

### 1. Validate

- Verify the audio file exists at the given path
- Verify the template file exists at `~/.claude/skills/transcribe/templates/<template_name>.docx`
- Determine output path: `<directory_of_mp3>/<mp3_name_without_ext>.docx`

### 2. Transcribe

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --transcribe '<mp3_path>' --output_dir /tmp
```

This saves `<mp3_name>.txt` in `/tmp`. Read the generated file.

### 3. Correct transcript

Review the raw transcript and fix speech recognition errors:
- Misrecognized words and homophones
- Names of people, products, and technical terms
- Missing or incorrect punctuation
- Sentence boundaries and paragraph breaks
- Filler words removal (uh, um, э, ну) unless they carry meaning

**Rules:**
- Fix ONLY recognition errors — do NOT change meaning, rephrase, or summarize
- Preserve the original speaker's wording as closely as possible
- If unsure whether something is an error, leave it as-is

Overwrite the `.txt` file in `/tmp` with the corrected version.

### 4. Extract template structure

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --extract '~/.claude/skills/transcribe/templates/<template_name>.docx'
```

This outputs the template structure showing all elements and placeholders.

### 5. Read style references (only if `--ref` flag is present)

**Skip this step if `--ref` was NOT provided.**

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --references '~/.claude/skills/transcribe/references/'
```

This outputs the text content of all .docx files in the reference directory. Use this as a **writing style guide** — match the tone, formality level, phrasing patterns, and structure from the reference documents when generating content.

### 6. Generate content JSON

Based on the transcript + template structure (+ style references if `--ref`), generate `content.json`:

```json
{
  "elements": [
    {
      "type": "table",
      "template_table": 0,
      "rows": [[{"text": "cell value"}]]
    },
    {
      "type": "paragraph",
      "style": "Heading 1",
      "text": "Section Title"
    },
    {
      "type": "paragraph",
      "style": "Normal",
      "runs": [
        {"text": "Bold part", "bold": true},
        {"text": " — italic part", "italic": true}
      ]
    },
    {
      "type": "paragraph",
      "style": "List Paragraph",
      "runs": [
        {"text": "Topic", "bold": true},
        {"text": " — description", "italic": true}
      ]
    }
  ]
}
```

**Rules:**
- Follow the template structure exactly: same order of element types
- Fill all `[placeholders]` with content from the transcript
- Add or remove repeating elements (headings, list items, table rows) as needed based on transcript content
- `template_table: N` — use formatting from the N-th table in the template (0-indexed)
- Use `runs` for mixed formatting (bold + italic in one paragraph)
- Use `text` shortcut when the entire paragraph has uniform style
- Use `\n` for line breaks within a single cell
- Keep the document language consistent with the template
- If `--ref` was used: **match the writing style from reference documents** (tone, vocabulary, sentence structure)

### 7. Save and build

Write the generated JSON to `/tmp/meeting_summary_<timestamp>.json`

Run:
```bash
python3 '~/.claude/skills/transcribe/process_docx.py' --fill '~/.claude/skills/transcribe/templates/<template_name>.docx' '/tmp/meeting_summary_<timestamp>.json' '<output_path>'
```

### 8. Cleanup

Delete temporary files:
- `/tmp/<mp3_name>.txt` (transcript)
- `/tmp/meeting_summary_<timestamp>.json` (content JSON)

### 9. Report

Tell the user the output file path.
