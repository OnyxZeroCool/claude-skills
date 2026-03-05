---
name: summarize
description: Use when user wants a summary of a file (txt, md, html, pdf) or web page, asks to "summarize", "explain this article", "what's in this file", or shares a URL for analysis. Triggers on "summarize", "саммари", "кратко", "пересказ", "tl;dr".
user_invocable: true
argument: "<path_or_url> [short|normal|detailed] [lang=XX]"
---

# Summarize

Generate structured summaries of files and web pages with three detail levels and follow-up Q&A.

## Usage

```
/summarize <path_or_url> [mode] [lang=XX]
```

- **Source:** local file path or URL
- **Mode:** `short`, `normal` (default), `detailed`
- **Language:** `lang=en`, `lang=de`, etc. Default: language of the source material

## Algorithm

### 1. Parse arguments

Extract source, mode, and language from the argument string. Default mode = `normal`.

**Determine source type:**
- Starts with `http://` or `https://` → URL
- Otherwise → file path (resolve relative to cwd)

### 2. Fetch content

**File:**
- Use the `Read` tool to read the file
- For PDF files >10 pages: read in chunks of 20 pages, concatenate content
- If file is empty or unreadable — inform the user and stop

**URL:**
- Use the `WebFetch` tool with prompt: "Extract all text content from this page. Return the full text without summarizing."
- If URL is unreachable — inform the user and stop

### 3. Generate summary

Use the **Agent tool** with `model: "sonnet"` and `subagent_type: "general-purpose"`.

Pass the fetched content and a prompt based on the mode and language. The agent must receive:
- The full content (or as much as fits)
- The mode template (see below)
- Language instruction: if `lang` is specified, generate summary in that language; otherwise, use the language of the source material

**Mode templates to include in the agent prompt:**

**`short` mode:**
```
Summarize in 3-5 sentences. Capture the core message and key takeaway.
Format: plain paragraph, no headers.
```

**`normal` mode:**
```
Create a structured summary:
## Summary
[5-7 sentence overview]

## Key Points
1. [Main point 1]
2. [Main point 2]
... (5-10 points)

## Conclusions
[Key takeaways and implications]
```

**`detailed` mode:**
```
Create a deep analysis:
## Summary
[5-7 sentence overview]

## Key Points
1. [Main point 1]
2. [Main point 2]
... (5-10 points)

## Detailed Breakdown
### [Topic/Section 1]
[Analysis with relevant quotes from the source]

### [Topic/Section 2]
[Analysis with relevant quotes from the source]
...

## Notable Quotes
> [Direct quotes worth highlighting]

## Critical Analysis
[Strengths, weaknesses, biases, missing perspectives]

## Conclusions
[Key takeaways and implications]
```

Output the agent's response to the user as-is.

### 4. Q&A mode

After the summary, print:

```
---
Source content is loaded. Ask any questions about this material.
```

The content remains in conversation context — answer follow-up questions using it.

## Edge Cases

- **Large PDF (>10 pages):** read in chunks of 20 pages, pass all to agent
- **Very large content:** warn user, truncate to ~15000 characters before sending to agent
- **Empty file:** inform user, stop
- **URL unreachable:** inform user, stop
- **Binary/unsupported file:** inform user, suggest converting to text first
