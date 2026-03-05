---
name: youtube
description: Use when user shares a YouTube link and wants a summary, transcript analysis, comment analysis, or Q&A about video content. Triggers on YouTube URLs, "summarize video", "video comments", "what's in this video".
user_invocable: true
argument: "<youtube_url> [summary|comments|full]"
---

# YouTube Video Analyzer

Fetch YouTube video transcript and/or comments, generate a structured summary, then answer follow-up questions.

## Usage

```
/youtube <url> [mode]
```

- **No mode or `summary`** — subtitles transcript + summary + Q&A
- **`comments`** — top comments analysis + Q&A
- **`full`** — transcript + comments combined analysis + Q&A

## Requirements

- `yt-dlp` (brew install yt-dlp)
- `jq` (brew install jq)

## Security

- All **URLs** and **file paths** in bash commands MUST be wrapped in single quotes to prevent shell injection
- Validate URL with strict regex: `^https?://(www\.)?(youtube\.com/(watch|shorts)|youtu\.be/)` — reject anything else
- Sanitize video_id before using in file paths: allow only `[A-Za-z0-9_-]`

## Authentication

YouTube requires browser cookies for most requests. To avoid multiple Keychain prompts, export cookies to a file **once** in step 2, then reuse the file for all subsequent yt-dlp calls:

- **Step 2 (metadata):** `--cookies-from-browser chrome --cookies '<WORK_DIR>/yt_<video_id>_cookies.txt'` — reads from Chrome, saves to file
- **All later steps:** `--cookies '<WORK_DIR>/yt_<video_id>_cookies.txt'` — reads from file, no browser access

## Algorithm

### 1. Parse arguments and resolve temp directory

Extract URL and mode from the argument string. Default mode = `summary`.

Validate URL with strict regex: `^https?://(www\.)?(youtube\.com/(watch|shorts)|youtu\.be/)`. Reject URLs that don't match.

**Resolve working directory** — run once in a **sandboxed** bash call:
```bash
echo $TMPDIR
```

Store the output as `WORK_DIR` (e.g. `/tmp/claude`). Use this **literal path** in ALL subsequent commands — never use the `$TMPDIR` variable, as it differs between sandbox and non-sandbox bash calls.

### 2. Fetch metadata (exports cookies)

This is the **only** call that touches the browser cookie store. It also saves cookies to a file for reuse.

```bash
yt-dlp --cookies-from-browser chrome --cookies '<WORK_DIR>/yt_<video_id>_cookies.txt' --dump-json --no-download '<url>' > '<WORK_DIR>/yt_<video_id>_meta.json'
```

Then extract and display with a single jq call:

```bash
jq -r '
  .duration as $d |
  ($d / 3600 | floor) as $h |
  (($d % 3600) / 60 | floor) as $m |
  ($d % 60 | floor) as $s |
  [
    "Title: \(.title // "?")",
    "Author: \(.uploader // "?")",
    "Date: \(.upload_date // "?")",
    if $h > 0
      then "Duration: \($h):\($m | tostring | if length < 2 then "0"+. else . end):\($s | tostring | if length < 2 then "0"+. else . end)"
      else "Duration: \($m):\($s | tostring | if length < 2 then "0"+. else . end)"
    end,
    "Description: \(.description // "" | .[0:500])"
  ] | join("\n")
' '<WORK_DIR>/yt_<video_id>_meta.json'
```

### 3. Fetch subtitles and/or comments

All calls below use `--cookies '<WORK_DIR>/yt_<video_id>_cookies.txt'` (NO `--cookies-from-browser`).

If mode requires both subtitles and comments (`full`), run them **in parallel**.

**Subtitles (summary / full modes):**
```bash
yt-dlp --cookies '<WORK_DIR>/yt_<video_id>_cookies.txt' --write-subs --write-auto-subs --sub-langs 'en' --skip-download --convert-subs srt -o '<WORK_DIR>/yt_%(id)s.%(ext)s' '<url>'
```

**Subtitle priority:** manual > auto-generated. Use `--sub-langs 'en'` for English; change to your preferred language code if needed.

Pick the best available file: `<WORK_DIR>/yt_<id>.en.srt`.

Strip SRT formatting and join into plain text with sed + uniq:

```bash
sed -E '/^[0-9]+$/d; /^[0-9]{2}:[0-9]{2}:[0-9]{2},[0-9]+ --> /d; s/<[^>]*>//g; /^[[:space:]]*$/d' '<WORK_DIR>/yt_<video_id>.<lang>.srt' | uniq | tr '\n' ' ' | head -c 15000
```

If NO subtitles are available — inform the user and suggest `comments` mode instead.

**Comments (comments / full modes):**
```bash
yt-dlp --cookies '<WORK_DIR>/yt_<video_id>_cookies.txt' --dump-json --write-comments --no-download --extractor-args 'youtube:comment_sort=top;max_comments=200,10,0' '<url>' > '<WORK_DIR>/yt_<video_id>_comments.json'
```

Extract top comments sorted by likes with jq:

```bash
jq -r '
  (.comments // []) |
  map(select(.parent == "root")) |
  sort_by(-.like_count) |
  .[0:25][] |
  "[\(.like_count // 0)👍] @\(.author // "?"): \(.text // "" | .[0:300])\n---"
' '<WORK_DIR>/yt_<video_id>_comments.json'
```

Check if comments are disabled:

```bash
jq '.comments | length' '<WORK_DIR>/yt_<video_id>_comments.json'
```

If result is 0 — inform the user.

### 4. Generate summary

**IMPORTANT:** Use the **Agent tool** with `model: "sonnet"` to generate the summary. Pass the transcript/comments text and metadata to the agent with a prompt to produce the structured markdown below. The agent's response is your summary — output it to the user as-is.

Structured markdown format:

**For `summary` mode:**

```
## [Video Title]
**Author:** ... | **Date:** ... | **Duration:** ...

### Summary
[5-7 sentence summary of the video content]

### Key Points
1. [Main point 1]
2. [Main point 2]
...

### Notable Quotes
> [Direct quote from transcript if noteworthy]
```

**For `comments` mode:**

```
## [Video Title] — Comment Analysis
**Author:** ... | **Comments:** [count]

### Top Themes
1. [Theme 1] — [brief description]
2. [Theme 2] — [brief description]

### Sentiment
[Overall sentiment: positive/negative/mixed, with brief explanation]

### Most Popular Comments
1. [comment text] — @author ([likes] likes)
...
```

**For `full` mode:** combine both sections above.

### 5. Enter Q&A mode

After the summary, print:

```
---
Transcript and metadata are loaded. Ask any questions about this video.
```

The transcript/comments remain in conversation context — answer follow-up questions using them.

### 6. Cleanup

After generating the summary, delete temporary files:

```bash
rm -f '<WORK_DIR>'/yt_<video_id>*
```

## Edge Cases

- **No subtitles available:** Tell user, suggest `comments` mode
- **Comments disabled:** Tell user, proceed with transcript only
- **Very long video (>2h):** Warn user that transcript is large, proceed normally
- **Shorts/clips:** Works the same way, just shorter content
- **Geo-restricted:** yt-dlp will fail — inform user about the restriction
