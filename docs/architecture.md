# Architecture

This document walks through how the automation hangs together end-to-end.

## High-level flow

```
┌──────────────┐    ┌───────────────────────┐    ┌────────────────────────┐
│   launchd    ├───►│   run_job_search.sh   ├───►│ extract_linkedin_urls  │
│  (3× daily)  │    │       (bash)          │    │  (osascript + Mail)    │
└──────────────┘    └──────────┬────────────┘    └─────────────┬──────────┘
                               │                               │ URLs
                               ▼                               ▼
                    ┌──────────────────────────────────────────────┐
                    │       Claude CLI agent (LLM brain)          │
                    │  ──────────────────────────────────────     │
                    │   • Indeed MCP                              │
                    │   • WebSearch / WebFetch                    │
                    │   • Reads dedup file                        │
                    │   • Applies hard filters                    │
                    │   • Scores matches 1–5                      │
                    └──────────────────────────┬───────────────────┘
                                               │
                          ┌────────────────────┼────────────────────┐
                          ▼                    ▼                    ▼
                   openpyxl writes        python-docx          appends URLs
                   to Excel tracker       generates resumes    to dedup file
                                               │
                                               ▼
                              ┌────────────────────────────────┐
                              │       send_daily_digest.py     │
                              │  ────────────────────────────  │
                              │   • matplotlib PNG masthead    │
                              │   • Refined-minimal HTML body  │
                              │   • Plain-text fallback        │
                              └────────────────┬───────────────┘
                                               ▼
                                  ┌──────────────────────────┐
                                  │   Mail.app via osascript │
                                  │   (multipart/alternative)│
                                  └────────────┬─────────────┘
                                               ▼
                                       Your inbox 📬
```

## Component breakdown

### `launchd` (macOS scheduler)
- Two LaunchAgents: daily (Mon–Fri × 3) and weekly (Friday 19:00)
- `StartCalendarInterval` for time-based firing
- `StandardOutPath` / `StandardErrorPath` go to `/tmp/` (TCC-clear)
- `EnvironmentVariables` block ensures `PATH`, `HOME`, `LANG` are set when
  invoked by `launchd` (which doesn't inherit a login shell's environment)

### `run_job_search.sh`
- `caffeinate -i -t 5400` keeps the Mac awake for the run (up to 90 min)
- `set -euo pipefail` for fail-fast semantics
- `trap` cleans up child processes and temp files on exit
- `sed` substitutes `{TODAY}` placeholder in the prompt
- All paths/credentials sourced from `.env` (never hardcoded)

### `extract_linkedin_urls.py`
- Uses Mail.app's `whose date received > cutoff and sender contains "linkedin"`
  predicate so filtering happens at the index level (millisecond-fast even
  with 9k+ inbox messages)
- Reads `source` (raw RFC822) instead of `content` to keep `href` attributes
  intact in the HTML body
- Decodes quoted-printable soft line breaks (`=\r\n`, `=\n`) which would
  otherwise split a URL across lines and truncate the job ID
- Normalizes every match to canonical `https://www.linkedin.com/jobs/view/{id}/`

### Claude CLI agent
- Invoked with `--print --dangerously-skip-permissions` for non-interactive
  execution under `launchd`
- The prompt enforces hard filters (recency ≤ 7 days, salary thresholds) so
  the LLM never has to argue about whether a job qualifies — it just counts
- Output schema is strict: `<<<EMAIL_START>>>` / `<<<EMAIL_END>>>` markers
  wrap a structured block the digest script can parse with regex, not NLP

### `send_daily_digest.py`
- **Composite PNG** (matplotlib Agg backend): masthead with date kicker,
  giant numeral, 3-column market KPIs with hairline-above styling, 7-day
  trend bar chart with today highlighted ink, filter footnote.
  Everything visual is rastered so iOS Mail's dark-mode auto-conversion
  cannot disturb it.
- **HTML body**: table-based layout with legacy `bgcolor` attributes plus
  inline CSS for max email-client safety. Refined-minimal aesthetic:
  hairline-separated rows, subtle score chips, monospace metadata, single
  ink accent.
- **Dark-mode lockdown**:
  - `<meta name="color-scheme" content="light only">`
  - `<meta name="supported-color-schemes" content="light">`
- **URL validation**: normalizes LinkedIn URLs to canonical form, rejects
  search/list/malformed URLs, drops the "View posting" button entirely
  when a URL fails validation rather than dead-ending on a broken link.

### Mail.app delivery (the MIME gotcha)
This is the most non-obvious part of the project.

When AppleScript sets only `html content` on a message AND adds an attachment,
the resulting MIME structure is:

```
multipart/mixed
├── text/html              ← Outlook iPhone treats this as ATT00001.HTM
└── application/...        ← Excel attachment
```

To get HTML rendered inline in every client (Mail.app, iOS Mail, Outlook
iPhone, Gmail), the script:

1. Sets `content` (plain text) FIRST
2. Sets `html content` LAST
3. Does NOT add an attachment

This produces:

```
multipart/alternative
├── text/plain
└── text/html              ← rendered inline everywhere
```

The Excel tracker is referenced by path in the footer instead of attached.

## Data flow

1. `extract_linkedin_urls.py` writes URLs to `/tmp/jobsearch_linkedin_urls.txt`
2. Claude CLI reads the prompt, sees the URLs file path in env, and uses
   `WebFetch` on each URL
3. Claude updates `$BASE_DIR/job_tracker.xlsx` via openpyxl
4. Claude creates per-job folders under `$BASE_DIR/Jobs/<Title> - <Company>/`
5. Claude appends new URLs (with today's date prefix) to `$BASE_DIR/found_jobs.txt`
6. Claude emits the structured email block on stdout
7. The shell captures stdout to `$TMP_OUTPUT` and passes it to `send_daily_digest.py`
8. The digest script parses the block, builds the PNG + HTML, sends via Mail.app
