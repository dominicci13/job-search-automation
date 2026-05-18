#!/bin/bash
# Daily job search runner — fired by launchd Mon-Fri at 06/12/18 local time.
#
# Loads .env from the repo root (one level above this script), then orchestrates:
#   1. Mail.app inbox URL extraction
#   2. Claude CLI agent run with the prompt template
#   3. Daily digest email send
#
# Designed to be idempotent: safe to re-run mid-day if a previous run failed.

set -euo pipefail

# ─── Locate this script's parent repo and load .env ───
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

if [ -f "$REPO_DIR/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  source "$REPO_DIR/.env"
  set +a
else
  echo "ERROR: $REPO_DIR/.env not found. Copy .env.example to .env and fill it in." >&2
  exit 1
fi

# ─── Required env vars (sourced from .env) ───
: "${BASE_DIR:?BASE_DIR not set in .env}"
: "${SCRIPTS_DIR:?SCRIPTS_DIR not set in .env}"
: "${EMAIL_TO:?EMAIL_TO not set in .env}"
: "${CLAUDE_BIN:?CLAUDE_BIN not set in .env}"
: "${PYTHON_BIN:?PYTHON_BIN not set in .env}"

CLAUDE_MODEL="${CLAUDE_MODEL:-claude-sonnet-4-6}"
PROMPT_FILE="$REPO_DIR/config/prompt.txt"
LOG_DIR="${LOG_DIR:-$BASE_DIR/logs}"
TODAY=$(date '+%Y-%m-%d')
TIMESTAMP=$(date '+%Y%m%d_%H%M')
LOG="$LOG_DIR/search_$TIMESTAMP.log"
TMP_OUTPUT="/tmp/job_search_output_$TIMESTAMP.txt"
LINKEDIN_URLS_FILE="/tmp/jobsearch_linkedin_urls.txt"

mkdir -p "$LOG_DIR" "$BASE_DIR"

# Ensure user identity and temp dir are set even when launched by launchd
export TMPDIR="${TMPDIR:-$(getconf DARWIN_USER_TEMP_DIR 2>/dev/null || echo /tmp/)}"
export USER="${USER:-$(whoami)}"
export LOGNAME="${LOGNAME:-$USER}"

# Keep the machine awake for the duration of this run (up to 90 minutes)
caffeinate -i -t 5400 &
CAFFEINATE_PID=$!
# shellcheck disable=SC2064
trap "kill $CAFFEINATE_PID 2>/dev/null; rm -f $TMP_OUTPUT" EXIT

echo "======================================" >> "$LOG"
echo "Job search started: $(date '+%Y-%m-%d %H:%M:%S %Z')" >> "$LOG"
echo "======================================" >> "$LOG"

# ─── Step 1: extract LinkedIn job URLs from inbox ───
# Non-fatal — falls through if Mail.app isn't reachable.
"$PYTHON_BIN" "$SCRIPTS_DIR/extract_linkedin_urls.py" \
    > "$LINKEDIN_URLS_FILE" 2>>"$LOG" || true
LINKEDIN_URL_COUNT=$(wc -l < "$LINKEDIN_URLS_FILE" | tr -d ' ')
echo "Extracted $LINKEDIN_URL_COUNT LinkedIn URLs from inbox" >> "$LOG"

# ─── Step 2: run the LLM agent with the prompt ───
# Capture the exit code without aborting the script — we want to send the
# digest even on failure so we get notified. `|| EXIT_CODE=$?` short-circuits
# `set -e` on non-zero exit; `EXIT_CODE=0` default ensures the var always
# exists for the log line below.
PROMPT=$(sed "s/{TODAY}/$TODAY/g" "$PROMPT_FILE")
EXIT_CODE=0
echo "$PROMPT" | "$CLAUDE_BIN" \
  --print \
  --dangerously-skip-permissions \
  --model "$CLAUDE_MODEL" \
  > "$TMP_OUTPUT" 2>&1 || EXIT_CODE=$?

cat "$TMP_OUTPUT" >> "$LOG"
echo "" >> "$LOG"
echo "Job search ended: $(date '+%Y-%m-%d %H:%M:%S %Z') (exit: $EXIT_CODE)" >> "$LOG"

# ─── Step 3: send the daily digest email ───
SUBJECT="Job Search — $TODAY $(date '+%H:%M')"
"$PYTHON_BIN" "$SCRIPTS_DIR/send_daily_digest.py" \
    "$TMP_OUTPUT" "$SUBJECT" "$EMAIL_TO" >> "$LOG" 2>&1

# Keep only the 30 most recent log files
ls -t "$LOG_DIR"/search_*.log 2>/dev/null | tail -n +31 | xargs rm -f 2>/dev/null

exit $EXIT_CODE
