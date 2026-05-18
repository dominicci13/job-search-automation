#!/bin/bash
# Weekly digest runner — fired by launchd every Friday at 19:00 local time.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

if [ -f "$REPO_DIR/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  source "$REPO_DIR/.env"
  set +a
else
  echo "ERROR: $REPO_DIR/.env not found." >&2
  exit 1
fi

: "${SCRIPTS_DIR:?SCRIPTS_DIR not set in .env}"
: "${PYTHON_BIN:?PYTHON_BIN not set in .env}"

LOG_DIR="${LOG_DIR:-$BASE_DIR/logs}"
LOG="$LOG_DIR/weekly_$(date '+%Y%m%d_%H%M').log"
mkdir -p "$LOG_DIR"

echo "Weekly digest started: $(date '+%Y-%m-%d %H:%M:%S %Z')" >> "$LOG"
"$PYTHON_BIN" "$SCRIPTS_DIR/weekly_digest.py" >> "$LOG" 2>&1
echo "Weekly digest ended: $(date '+%Y-%m-%d %H:%M:%S %Z') (exit: $?)" >> "$LOG"
