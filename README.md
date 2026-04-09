# Company Researcher

Google Apps Script that enriches a list of company names with configurable data points using the Claude API with web search. Results and confidence scores are written back to the same Google Sheet.

## Setup

1. Open your Google Sheet
2. Create two tabs:
   - **`Companies`** — add company names in column A starting at row 2
   - **`Research Fields`** — define what to research (see below)
3. Open Extensions → Apps Script
4. Copy all `.gs` files and `appsscript.json` into the project
5. Set your API key: Project Settings → Script Properties → Add `ANTHROPIC_API_KEY`
   (Or skip this step — you'll be prompted the first time you run)
6. Reload the spreadsheet — the **Company Research** menu will appear

## Research Fields Tab

| Column A | Column B | Column C |
|----------|----------|----------|
| Field name | Instructions for Claude | Expected format |
| `HQ City` | `Return city and state only, e.g. "Austin, TX"` | `text` |
| `Website` | `Official website URL` | `URL` |
| `For-Profit or Non-Profit` | `Return "For-Profit" or "Non-Profit"` | `text` |

Column B and C are optional but improve accuracy and consistency.

## Usage

- **Company Research → Run Research** — processes all companies without results. Safe to re-run; already-processed rows are skipped. If the script times out (Apps Script 6-minute limit), re-run and it resumes from where it left off.
- **Company Research → Clear Results** — removes all generated columns. Company names in column A are preserved.

## Output

For each configured field, two columns are written next to each other:
- **Value** — Claude's answer
- **Confidence** (0–100) — color-coded: green 80–100, yellow 50–79, red 0–49

A **Notes** column (last column) is written when any field confidence is below 70, with a one-line note explaining the uncertainty.

## Configuration

Edit constants at the top of `Config.gs`:
- `API_CALL_DELAY_MS` (default: 1000) — delay between companies in ms; increase if you hit rate limits
- `CONFIDENCE_NOTES_THRESHOLD` (default: 70) — confidence below this writes a note

## Notes

- Requires an Anthropic API key with access to `claude-sonnet-4-6` and the `web_search` tool (`anthropic-beta: web-search-2025-03-05`).
- Each company is one Claude API call. For 100 companies at 1-second delay, expect 5–7 minutes.
- For 200+ companies, the script may hit Apps Script's 6-minute execution limit — just re-run and it picks up where it left off.
