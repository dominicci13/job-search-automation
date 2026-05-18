#!/usr/bin/env python3
"""Extract LinkedIn job-posting URLs from recent inbox emails.

Uses Mail.app's `whose` predicate to filter at index level (millisecond-fast)
rather than iterating all inbox messages. Targets emails from linkedin.com
senders in the last 14 days, pulls canonical
https://www.linkedin.com/jobs/view/<id>/ URLs, prints one per line.

Output is capped at MAX_URLS to bound WebFetch cost.
"""
import re
import subprocess
import sys

MAX_URLS = 60
LOOKBACK_DAYS = 14

OSA = r'''
tell application "Mail"
  set cutoff to (current date) - (LOOKBACK * days)
  set msgs to (messages of inbox whose date received > cutoff and sender contains "linkedin")
  set out to ""
  repeat with m in msgs
    try
      -- Use `source` to get the raw RFC822 (headers + HTML body) so href URLs
      -- are visible. `content` returns rendered plain text, which strips hrefs.
      set out to out & (source of m) & (return & "===EMAIL_SEP===" & return)
    end try
  end repeat
  return out
end tell
'''.replace("LOOKBACK", str(LOOKBACK_DAYS))


def main():
    try:
        result = subprocess.run(
            ["osascript", "-e", OSA],
            capture_output=True,
            text=True,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        print("extract_linkedin_urls: osascript timed out (120s)", file=sys.stderr)
        return 0
    if result.returncode != 0:
        print(f"extract_linkedin_urls: osascript error: {result.stderr.strip()}", file=sys.stderr)
        return 0

    # Decode quoted-printable soft line breaks (=\r\n or =\n) which split URLs
    # across lines in the raw RFC822 source. Without this, the URL digits get
    # truncated at the line wrap (every ~76 chars) and we collect short IDs
    # like "4414" instead of "4414254639".
    text = result.stdout.replace("=\r\n", "").replace("=\n", "")

    # LinkedIn job IDs are 8-12 digits; reject anything shorter as a truncation.
    pat = re.compile(
        r"https?://[a-z]{0,3}\.?linkedin\.com/(?:jobs/view|comm/jobs/view)/(\d{8,12})",
        re.IGNORECASE,
    )
    ids_seen = set()
    out_urls = []
    for match in pat.finditer(text):
        job_id = match.group(1)
        if job_id in ids_seen:
            continue
        ids_seen.add(job_id)
        out_urls.append(f"https://www.linkedin.com/jobs/view/{job_id}/")
        if len(out_urls) >= MAX_URLS:
            break

    for u in out_urls:
        print(u)
    print(
        f"extract_linkedin_urls: {len(out_urls)} unique LinkedIn URLs from last {LOOKBACK_DAYS} days",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
