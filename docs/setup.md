# Setup

End-to-end install on macOS. Should take ~15 minutes.

## 1. Prerequisites

- macOS (tested on Sonoma & Sequoia)
- Python 3.11+
- [Claude CLI](https://docs.claude.com/en/docs/claude-code/overview) installed and authenticated
- Mail.app configured with an email account (iCloud, Gmail, etc. — anything works)

## 2. Clone and configure

```bash
git clone https://github.com/dominicci13/job-search-automation.git
cd job-search-automation

# Copy the example configs and edit them
cp .env.example .env
cp config/profile.example.yaml config/profile.yaml
cp config/prompt.template.txt config/prompt.txt
```

Edit `.env` with your paths and email address.

Edit `config/profile.yaml` with your name, contact, skills, target markets, and
visa strategy (or remove the visa block if irrelevant).

Edit `config/prompt.txt` to customize the search queries to your stack.

## 3. Install Python dependencies

```bash
pip install -r requirements.txt
```

## 4. Create the blank Excel tracker

```bash
python scripts/setup_excel.py
```

This writes `$BASE_DIR/job_tracker.xlsx` with the 29-column schema, dropdowns,
freeze panes, and autofilter.

## 5. Place your base resume

Put a `resume_base.docx` file in `$BASE_DIR/`. This is what the per-job
resume generator uses as the starting point. Keep a clean, ATS-friendly
version — the agent tailors it per role.

## 6. Test a dry run

```bash
bash scripts/run_job_search.sh
```

Logs go to `$LOG_DIR/search_<timestamp>.log`. Check that:

- The Claude CLI invocation succeeds
- The Excel tracker gets new rows
- An email arrives in your inbox

## 7. Install the launchd agents

Copy the example plists, edit the paths, and load them:

```bash
# Daily — runs Mon-Fri at 06:00, 12:00, 18:00
cp launchd/com.example.jobsearch.plist.example \
   ~/Library/LaunchAgents/com.YOU.jobsearch.plist

# Open it and replace all /REPLACE_ME placeholders with your actual paths
open -e ~/Library/LaunchAgents/com.YOU.jobsearch.plist

launchctl load ~/Library/LaunchAgents/com.YOU.jobsearch.plist

# Weekly — runs Friday at 19:00
cp launchd/com.example.jobsearch-weekly.plist.example \
   ~/Library/LaunchAgents/com.YOU.jobsearch-weekly.plist

open -e ~/Library/LaunchAgents/com.YOU.jobsearch-weekly.plist

launchctl load ~/Library/LaunchAgents/com.YOU.jobsearch-weekly.plist
```

Verify the agents are scheduled:

```bash
launchctl list | grep jobsearch
```

You should see both agents listed with a recent or upcoming run time.

## 8. (Optional) Wake the Mac for the 6 AM run

If your Mac is normally asleep at 6 AM, schedule a wake event:

```bash
sudo pmset repeat wakeorpoweron MTWRF 05:55:00
```

This wakes (or powers on) the Mac at 5:55 AM Mon–Fri, 5 minutes before the
first scheduled run.

## 9. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| No email arrives | TCC permissions block Mail.app/Python from `osascript` | Run the script once manually; macOS will prompt for permissions. Approve them. |
| Empty body in Outlook iPhone | Attachment is forcing `multipart/mixed` | The daily template intentionally has no attachment — verify your customization didn't add one back |
| `claude: command not found` under launchd | `$PATH` not set in the plist | Hardcode `CLAUDE_BIN` in `.env` as the absolute path |
| `extract_linkedin_urls.py` times out | Mail.app indexing thousands of messages | The script uses a `whose` predicate to filter at index time — should be fast. If still slow, reduce `LOOKBACK_DAYS` |
| LLM agent returns shorter than expected | Context limit hit, or model rate-limited | Tighten the prompt or switch model in `.env` |
