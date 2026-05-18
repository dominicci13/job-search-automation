#!/usr/bin/env python3
"""Send the daily Job Search Summary email as HTML via Mail.app.

Strategy: bake the design-heavy parts (header, KPI tiles, chart) into a single
composited PNG image. Mail.app's dark-mode auto-conversion strips gradient
backgrounds and card fills from HTML divs — but it cannot strip styling from a
raster image. HTML around the image stays minimal: just job cards (so links
work) and a footer.

Result: the daily renders identically in Mail.app, iOS Mail, and Outlook
regardless of dark-mode override behavior.
"""
import base64
import datetime
import io
import os
import re
import subprocess
import sys
import tempfile
from html import escape

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import gridspec, rcParams


def _require_env(name):
    val = os.environ.get(name)
    if not val:
        raise SystemExit(
            f"{name} not set — source your .env before running (e.g. via run_job_search.sh)"
        )
    return val


BASE_DIR = _require_env("BASE_DIR")
EXCEL_PATH = os.path.join(BASE_DIR, "job_tracker.xlsx")
HISTORICAL_LOGS_DIR = os.environ.get("LOG_DIR") or os.path.join(BASE_DIR, "logs")

C = {
    "paper":        "#FFFFFF",
    "paper_warm":   "#FAFAFA",
    "ink":          "#0A0A0A",
    "ink_soft":     "#262626",
    "muted":        "#737373",
    "subtle":       "#A3A3A3",
    "hair":         "#E5E5E5",
    "hair_strong":  "#D4D4D4",
    "accent":       "#1F2937",
    "accent_soft":  "#F3F4F6",
    "score_high":   "#0F766E",
    "score_high_bg":"#F0FDFA",
    "score_mid":    "#0369A1",
    "score_mid_bg": "#F0F9FF",
    "score_low":    "#737373",
    "score_low_bg": "#FAFAFA",
    "bg":           "#F5F5F5",
    "card":         "#FFFFFF",
    "border":       "#E5E7EB",
    "gray_soft":    "#F5F5F5",
    "primary":      "#0A0A0A",
    "primary_soft": "#F3F4F6",
    "primary_dark": "#0A0A0A",
    "green":        "#0F766E",
    "green_soft":   "#F0FDFA",
    "gray":         "#737373",
    "spain":        "#0A0A0A",
    "us":           "#0A0A0A",
    "latam":        "#0A0A0A",
}

MARKET_EMOJI = {"Spain": "🇪🇸", "US Remote": "🇺🇸", "LATAM Agency": "🌎"}


# ──────────────────────────────────────────────────────────────────────
# Parsing (unchanged)
# ──────────────────────────────────────────────────────────────────────


def extract_email_block(content):
    s = content.find("<<<EMAIL_START>>>")
    e = content.find("<<<EMAIL_END>>>")
    if s != -1 and e != -1 and e > s:
        return content[s + len("<<<EMAIL_START>>>"):e].strip()
    if s != -1:
        return content[s + len("<<<EMAIL_START>>>"):].strip()
    return content.strip()


def parse_summary(block):
    keys = [
        "Spain jobs found", "US Remote jobs found", "LATAM jobs found",
        "Total new jobs", "Folders created",
        "Jobs skipped (dup)", "Jobs skipped (old)", "Jobs skipped (filter)",
    ]
    out = {}
    for k in keys:
        m = re.search(re.escape(k) + r":\s*(.+?)(?:\n|$)", block)
        if m:
            out[k] = m.group(1).strip()
    return out


def parse_jobs(block):
    idx = block.find("NEW JOBS DETAIL")
    if idx == -1:
        return []
    detail = block[idx:]
    raw_blocks = re.split(r"\n(?=\[SCORE:)", detail)
    jobs = []
    for rb in raw_blocks:
        rb = rb.strip()
        if not rb.startswith("[SCORE:"):
            continue
        lines = [ln.strip() for ln in rb.splitlines() if ln.strip()]
        if not lines:
            continue
        head = re.match(r"\[SCORE:\s*(\d)/5\]\s*(.+?)\s*@\s*(.+)", lines[0])
        if not head:
            continue
        job = {
            "score": int(head.group(1)), "title": head.group(2).strip(),
            "company": head.group(3).strip(),
            "market": "", "location": "", "mode": "", "salary": "",
            "posted": "", "deadline": "", "board": "", "url": "", "why": "",
        }
        for ln in lines[1:]:
            if ln.startswith("Why it matches"):
                _, _, val = ln.partition(":")
                job["why"] = val.strip()
                continue
            if ln.startswith("URL:"):
                job["url"] = ln[len("URL:"):].strip()
                continue
            for chunk in ln.split("|"):
                if ":" in chunk:
                    k, _, v = chunk.partition(":")
                    job[k.strip().lower().replace(" ", "_")] = v.strip()
        jobs.append(job)
    return jobs


def collect_7day_history():
    today = datetime.date.today()
    by_day = {}
    for i in range(7):
        d = today - datetime.timedelta(days=6 - i)
        by_day[d] = 0
    if not os.path.isdir(HISTORICAL_LOGS_DIR):
        return [(d, by_day[d]) for d in sorted(by_day.keys())]
    for fname in os.listdir(HISTORICAL_LOGS_DIR):
        m = re.match(r"search_(\d{4})(\d{2})(\d{2})_\d+\.log", fname)
        if not m:
            continue
        try:
            d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        if d not in by_day:
            continue
        try:
            with open(os.path.join(HISTORICAL_LOGS_DIR, fname), "r", errors="replace") as f:
                mm = re.search(r"Total new jobs:\s*(\d+)", f.read())
            if mm:
                by_day[d] = max(by_day[d], int(mm.group(1)))
        except OSError:
            continue
    return [(d, by_day[d]) for d in sorted(by_day.keys())]


# ──────────────────────────────────────────────────────────────────────
# Composite image — everything visual baked into one PNG
# ──────────────────────────────────────────────────────────────────────


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))


def build_composite_image(today_dt, total_new, summary, history_7d):
    """Refined minimal masthead — pure white, ink-on-white sans-serif typography,
    single muted accent, hairline dividers, generous whitespace. The visual
    hierarchy is carried entirely by typography weight and scale rather than
    color or decoration. Renders as a single PNG to lock down dark-mode
    behavior across all mail clients."""
    sans_stack  = ["SF Pro Display", "Helvetica Neue", "Helvetica",
                   "Arial", "DejaVu Sans"]
    rcParams["font.family"] = sans_stack
    rcParams["font.size"] = 11
    rcParams["font.weight"] = "regular"

    fig = plt.figure(figsize=(12, 11.5), facecolor=C["paper"])
    gs = gridspec.GridSpec(
        4, 3, figure=fig,
        height_ratios=[3.2, 1.6, 2.6, 0.5],
        hspace=0.70, wspace=0.30,
        left=0.08, right=0.92, top=0.96, bottom=0.04,
    )

    # ─── Section A: Quiet masthead — kicker → hero numeral → subtitle ───
    ax_header = fig.add_subplot(gs[0, :])
    ax_header.set_facecolor(C["paper"])
    ax_header.set_xlim(0, 1); ax_header.set_ylim(0, 1); ax_header.axis("off")

    # Top date kicker — small, tracked-out, muted
    weekday_str = today_dt.strftime("%a, %b %-d, %Y").upper()
    ax_header.text(0.0, 0.94, weekday_str,
                   color=C["muted"], fontsize=10, fontweight="semibold",
                   ha="left", va="center", family=sans_stack,
                   transform=ax_header.transAxes)

    # Top-right ID label
    ax_header.text(1.0, 0.94, "DAILY · JOB SEARCH",
                   color=C["muted"], fontsize=10, fontweight="semibold",
                   ha="right", va="center", family=sans_stack,
                   transform=ax_header.transAxes)

    # Hairline divider beneath the kicker
    ax_header.plot([0, 1], [0.86, 0.86], color=C["hair"], linewidth=0.8,
                   transform=ax_header.transAxes)

    # Giant ink numeral — left-aligned, generous space around it
    ax_header.text(0.0, 0.42, str(total_new),
                   color=C["ink"], fontsize=148, fontweight="bold",
                   ha="left", va="center", family=sans_stack,
                   transform=ax_header.transAxes)

    # Subtitle line, sized to read at a glance
    plural = "matches" if total_new != 1 else "match"
    ax_header.text(0.0, 0.06, f"new {plural} this run",
                   color=C["muted"], fontsize=18, fontweight="regular",
                   ha="left", va="center", family=sans_stack,
                   transform=ax_header.transAxes)

    # ─── Section B: 3 KPI cells — minimal, no borders, hairline above ───
    spain_n = int(summary.get("Spain jobs found", "0") or 0)
    us_n    = int(summary.get("US Remote jobs found", "0") or 0)
    latam_n = int(summary.get("LATAM jobs found", "0") or 0)
    kpi_data = [
        ("Spain",        spain_n, "Madrid · HQP / DNV"),
        ("United States", us_n,   "Remote · DNV"),
        ("LATAM",        latam_n, "Nearshore · DNV"),
    ]
    for i, (label, value, sub) in enumerate(kpi_data):
        ax = fig.add_subplot(gs[1, i])
        ax.set_facecolor(C["paper"])
        ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
        # hairline above
        ax.plot([0, 0.85], [0.95, 0.95], color=C["hair_strong"],
                linewidth=1.0, transform=ax.transAxes)
        # category label — small, tracked, muted
        ax.text(0.0, 0.78, label.upper(), color=C["muted"], fontsize=10,
                fontweight="semibold", ha="left", va="center",
                family=sans_stack, transform=ax.transAxes)
        # the value — quiet, ink
        num_color = C["ink"] if value > 0 else C["subtle"]
        ax.text(0.0, 0.42, str(value), color=num_color, fontsize=54,
                fontweight="bold", ha="left", va="center",
                family=sans_stack, transform=ax.transAxes)
        # subtitle in muted regular
        ax.text(0.0, 0.05, sub, color=C["muted"], fontsize=11,
                ha="left", va="center", family=sans_stack,
                transform=ax.transAxes)

    # ─── Section C: 7-day trend — quiet bar chart, ink on white ───
    ax_trend = fig.add_subplot(gs[2, :])
    ax_trend.set_facecolor(C["paper"])
    if history_7d:
        dates = [d for d, _ in history_7d]
        vals = [v for _, v in history_7d]
        labels = [d.strftime("%a\n%-m/%-d") for d in dates]
        # Today: ink. Others: hair_strong. Quiet contrast.
        bar_colors = [
            C["ink"] if i == len(vals) - 1 else C["hair_strong"]
            for i in range(len(vals))
        ]
        bars = ax_trend.bar(range(len(vals)), vals, color=bar_colors,
                            edgecolor=C["paper"], linewidth=0, width=0.56)
        ax_trend.set_xticks(range(len(labels)))
        ax_trend.set_xticklabels(labels, fontsize=10, color=C["muted"],
                                 family=sans_stack)
        ax_trend.set_title("LAST 7 DAYS",
                           loc="left", fontsize=10, fontweight="semibold",
                           color=C["muted"], pad=20, family=sans_stack)
        for spine in ("top", "right", "left", "bottom"):
            ax_trend.spines[spine].set_visible(False)
        ax_trend.tick_params(left=False, labelleft=False, bottom=False,
                             pad=10)
        ax_trend.grid(False)
        if any(vals):
            ax_trend.set_ylim(0, max(vals) + 2)
        for b, v in zip(bars, vals, strict=True):
            if v > 0:
                ax_trend.text(
                    b.get_x() + b.get_width()/2, v + 0.18, str(v),
                    ha="center", va="bottom", fontsize=12,
                    fontweight="semibold", color=C["ink"],
                    family=sans_stack,
                )

    # ─── Section D: Filter footnote — quiet, single line, no rules ───
    ax_filt = fig.add_subplot(gs[3, :])
    ax_filt.set_facecolor(C["paper"])
    ax_filt.set_xlim(0, 1); ax_filt.set_ylim(0, 1); ax_filt.axis("off")
    dup_n  = summary.get("Jobs skipped (dup)", "0")
    old_n  = summary.get("Jobs skipped (old)", "0")
    filt_n = summary.get("Jobs skipped (filter)", "0")
    # single thin hairline above
    ax_filt.plot([0, 1], [0.85, 0.85], color=C["hair"], linewidth=0.8,
                 transform=ax_filt.transAxes)
    ax_filt.text(
        0.0, 0.35,
        f"Filtered  ·  {dup_n} duplicate  ·  {old_n} older than 7 days  "
        f"·  {filt_n} salary or authorisation",
        color=C["muted"], fontsize=11, fontweight="regular",
        ha="left", va="center", family=sans_stack,
        transform=ax_filt.transAxes,
    )

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, facecolor=C["paper"], bbox_inches=None)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


# ──────────────────────────────────────────────────────────────────────
# Job cards — kept in HTML so links remain clickable
# ──────────────────────────────────────────────────────────────────────


def market_emoji(m): return MARKET_EMOJI.get(m, "📍")


def _safe_int(s, default=0):
    try: return int(str(s).strip().split()[0])
    except (ValueError, TypeError): return default


SANS = ("-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', "
        "'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif")
MONO = ("ui-monospace, 'SF Mono', Menlo, Monaco, 'Cascadia Mono', Consolas, "
        "'Liberation Mono', monospace")


# ──────────────────────────────────────────────────────────────────────
# URL validation — strip, normalize, reject anything that doesn't look like
# a real job posting URL. The 'READ THE POSTING' button is suppressed when
# a URL fails validation so the user never clicks through to a broken or
# unrelated link.
# ──────────────────────────────────────────────────────────────────────

_LINKEDIN_ID_RE = re.compile(
    r"linkedin\.com/(?:comm/)?jobs/view/(\d{8,12})", re.IGNORECASE
)
_GENERIC_URL_RE = re.compile(r"^https?://[A-Za-z0-9.\-]+(?:/[^\s]*)?$")


def validate_and_normalize_url(raw):
    """Return a clean URL string if `raw` looks like a real job-posting URL,
    else None. LinkedIn URLs are normalized to the canonical
    https://www.linkedin.com/jobs/view/{id}/ form. Search URLs, malformed
    URLs, and obvious junk are rejected."""
    if not raw:
        return None
    u = str(raw).strip().strip(",.;:)('\"<>")
    if not u or u.lower() in {"n/a", "none", "tbd", "tba", "unknown"}:
        return None

    # LinkedIn: extract job ID and rebuild canonical URL — this collapses all
    # variants (/comm/jobs/view, query strings, regional subdomains, etc.)
    m = _LINKEDIN_ID_RE.search(u)
    if m:
        return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"

    # Reject obvious non-postings: search pages, list pages
    lowered = u.lower()
    for bad in ("/jobs/search", "/search?", "/jobs?", "google.com/search",
                "linkedin.com/jobs/collections", "indeed.com/jobs?",
                "infojobs.net/ofertas-trabajo", "/results?"):
        if bad in lowered:
            return None

    # Generic URL — must start with http(s):// and parse cleanly
    if not _GENERIC_URL_RE.match(u):
        return None
    return u


def score_chip(score):
    """Subtle score pill — soft tinted background, never loud."""
    score = max(0, min(5, int(score or 0)))
    if score >= 4:
        bg, fg = C["score_high_bg"], C["score_high"]
    elif score == 3:
        bg, fg = C["score_mid_bg"], C["score_mid"]
    else:
        bg, fg = C["score_low_bg"], C["score_low"]
    return (
        f'<span style="display:inline-block;padding:3px 9px;border-radius:10px;'
        f'background:{bg};color:{fg};font-size:11px;font-weight:600;'
        f'letter-spacing:0.2px;white-space:nowrap;border:1px solid {C["hair"]};">'
        f'{score} / 5</span>'
    )


def market_label(market):
    """Plain market label, lower-key (no caps screaming)."""
    if not market:
        return "Market — TBD"
    return market.strip()


def job_card_html(job, number):
    """Minimal job entry — clean row with subtle hairline separator. Title +
    metadata + score chip on the same visual plane. Why-block is plain quiet
    text (no quotes, no italic flourishes). Link is a quiet text link in ink
    with subtle hover-style underline."""
    score = job.get("score", 0)
    market = job.get("market", "")
    title = escape(job.get("title", ""))[:130]
    company = escape(job.get("company", ""))[:90]
    location = escape(job.get("location", ""))[:80]
    mode = escape(job.get("mode", ""))
    salary = escape(job.get("salary", ""))
    posted = escape(job.get("posted", ""))
    board = escape(job.get("board", ""))
    url = validate_and_normalize_url(job.get("url", ""))
    why = escape(job.get("why", "")).strip()

    # Tertiary meta line — company, location, mode, salary
    parts = []
    if company:  parts.append(f'<span style="color:{C["ink"]};font-weight:600;">{company}</span>')
    if location: parts.append(location)
    if mode:     parts.append(mode)
    if salary and "not listed" not in salary.lower():
        parts.append(salary)
    facts_html = '<span style="color:' + C["subtle"] + ';"> · </span>'.join(parts)

    # Quaternary line — market + posted/board
    sub_parts = []
    if market:   sub_parts.append(market_label(market))
    if posted:   sub_parts.append(posted)
    if board:    sub_parts.append(board)
    sub_line = '  ·  '.join(sub_parts)

    # Why — plain quiet paragraph, no quotes
    why_block = ""
    if why:
        why_block = f"""
        <p style="font-family:{SANS};font-size:13px;line-height:1.65;
                  color:{C['ink_soft']};margin:12px 0 0 0;">
          {why}
        </p>
        """

    # Action link — minimal text link in ink with arrow
    if url:
        btn_html = f"""
        <a href="{escape(url, quote=True)}" target="_blank"
           style="display:inline-block;font-family:{SANS};font-size:13px;
                  color:{C['ink']};text-decoration:none;font-weight:600;
                  margin-top:14px;border-bottom:1px solid {C['ink']};
                  padding-bottom:1px;">View posting&nbsp;→</a>
        """
    else:
        btn_html = f"""
        <p style="font-family:{SANS};font-size:12px;color:{C['subtle']};
                  margin:14px 0 0 0;">
          Link unavailable — search "{title} {company}" on the source board.
        </p>
        """

    return f"""
    <table role="presentation" cellpadding="0" cellspacing="0" border="0"
           width="100%" style="margin:0;border-top:1px solid {C['hair']};">
      <tr>
        <td style="padding:24px 0 28px 0;">
          <!-- Top row: number/sub on left, score chip on right -->
          <table cellpadding="0" cellspacing="0" border="0" width="100%"
                 style="margin:0 0 10px 0;">
            <tr>
              <td valign="top" style="font-family:{MONO};font-size:11px;
                  color:{C['subtle']};letter-spacing:0.5px;">
                {number:02d} &nbsp;·&nbsp; {sub_line}
              </td>
              <td valign="top" align="right" style="white-space:nowrap;">
                {score_chip(score)}
              </td>
            </tr>
          </table>
          <!-- Title -->
          <h2 style="font-family:{SANS};font-size:20px;font-weight:600;
                     color:{C['ink']};margin:0 0 6px 0;line-height:1.3;
                     letter-spacing:-0.2px;">{title}</h2>
          <!-- Facts -->
          <p style="font-family:{SANS};font-size:14px;color:{C['ink_soft']};
                    margin:0;line-height:1.5;">
            {facts_html}
          </p>
          {why_block}
          {btn_html}
        </td>
      </tr>
    </table>
    """


def build_html(summary, jobs, today_dt, composite_b64):
    high_score_jobs = [j for j in jobs if j.get("score", 0) >= 4]
    other_jobs = [j for j in jobs if j.get("score", 0) < 4]

    def section_heading(label, count=None):
        count_html = ""
        if count is not None:
            count_html = (
                f'<span style="font-family:{MONO};font-size:12px;'
                f'color:{C["subtle"]};font-weight:500;margin-left:10px;">{count}</span>'
            )
        return f"""
        <div style="margin:36px 0 4px 0;">
          <h2 style="font-family:{SANS};font-size:13px;font-weight:600;
                     color:{C['muted']};margin:0;text-transform:uppercase;
                     letter-spacing:1.4px;">{label}{count_html}</h2>
        </div>
        """

    cards_html = ""
    counter = 1
    if high_score_jobs:
        cards_html += section_heading("Strong matches", count=len(high_score_jobs))
        for j in high_score_jobs:
            cards_html += job_card_html(j, counter)
            counter += 1
        if other_jobs:
            cards_html += section_heading("Other postings", count=len(other_jobs))
            for j in other_jobs:
                cards_html += job_card_html(j, counter)
                counter += 1
    elif jobs:
        cards_html += section_heading("New postings", count=len(jobs))
        for j in jobs:
            cards_html += job_card_html(j, counter)
            counter += 1
    else:
        cards_html = f"""
        <div style="padding:32px 0;border-top:1px solid {C['hair']};
                    border-bottom:1px solid {C['hair']};text-align:center;">
          <p style="font-family:{SANS};font-size:15px;color:{C['ink_soft']};
                    margin:0;">
            No new postings cleared the filters this run.
          </p>
          <p style="font-family:{SANS};font-size:12px;color:{C['muted']};
                    margin:6px 0 0 0;">
            See filtered counts in the header.
          </p>
        </div>
        """

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <meta name="color-scheme" content="light only">
  <meta name="supported-color-schemes" content="light">
  <meta name="format-detection" content="telephone=no,date=no,address=no,email=no,url=no">
  <title>Daily — Job Search</title>
  <style>
    :root {{ color-scheme: light only; supported-color-schemes: light; }}
    body {{ margin:0; padding:0; }}
    a {{ color: {C['ink']}; }}
  </style>
</head>
<body bgcolor="{C['bg']}" style="margin:0;padding:0;background-color:{C['bg']};
       font-family:{SANS};color:{C['ink']};">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
         bgcolor="{C['bg']}" style="background-color:{C['bg']};">
    <tr><td align="center" style="padding:40px 16px;">

      <table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0"
             bgcolor="{C['paper']}" style="background-color:{C['paper']};
             max-width:640px;width:100%;border:1px solid {C['hair']};
             border-radius:6px;">

        <!-- Masthead PNG -->
        <tr><td bgcolor="{C['paper']}" style="background-color:{C['paper']};
            padding:0;line-height:0;font-size:0;border-radius:6px 6px 0 0;">
          <img src="data:image/png;base64,{composite_b64}"
               width="640" alt="Daily Job Search"
               style="display:block;width:100%;max-width:640px;height:auto;
                      border:0;outline:none;border-radius:6px 6px 0 0;" />
        </td></tr>

        <!-- Body -->
        <tr><td bgcolor="{C['paper']}" style="background-color:{C['paper']};
            padding:8px 40px 32px 40px;">
          {cards_html}
        </td></tr>

        <!-- Footer -->
        <tr><td bgcolor="{C['paper_warm']}" style="background-color:{C['paper_warm']};
            padding:24px 40px;border-top:1px solid {C['hair']};
            border-radius:0 0 6px 6px;">
          <table cellpadding="0" cellspacing="0" border="0" width="100%">
            <tr>
              <td valign="top" style="font-family:{SANS};font-size:12px;
                  color:{C['muted']};line-height:1.7;">
                Generated {today_dt.strftime("%b %-d, %Y")}<br>
                Tracker: <span style="font-family:{MONO};color:{C['ink_soft']};
                  font-size:11px;">{escape(EXCEL_PATH)}</span>
              </td>
              <td valign="top" align="right" style="font-family:{SANS};font-size:11px;
                  color:{C['subtle']};font-weight:500;letter-spacing:0.5px;
                  text-transform:uppercase;">
                Daily
              </td>
            </tr>
          </table>
        </td></tr>

      </table>

    </td></tr>
  </table>
</body>
</html>"""


# ──────────────────────────────────────────────────────────────────────
# Send
# ──────────────────────────────────────────────────────────────────────


def build_plain_text_fallback(summary, jobs, today_dt):
    """Minimal plain-text alternative. Mail.app needs `content` set BEFORE
    `html content` to build a proper multipart/alternative MIME structure;
    without it, Outlook iPhone treats the HTML part as an attachment and
    shows a blank body."""
    total = summary.get("Total new jobs", "?")
    lines = [
        "THE DAILY — JOB SEARCH BRIEF",
        f"{today_dt.strftime('%A, %B %d, %Y')}",
        "=" * 50,
        f"{total} new matches on the wire",
        "",
        f"Spain: {summary.get('Spain jobs found', '0')}   "
        f"US Remote: {summary.get('US Remote jobs found', '0')}   "
        f"LATAM: {summary.get('LATAM jobs found', '0')}",
        "",
        "View this brief in HTML for the full layout.",
        "",
    ]
    for i, j in enumerate(jobs, 1):
        url = validate_and_normalize_url(j.get("url", "")) or ""
        lines.append(f"{i:02d}. [★{j.get('score', 0)}] {j.get('title', '')}")
        lines.append(f"    @ {j.get('company', '')} · {j.get('location', '')}")
        if url:
            lines.append(f"    {url}")
        lines.append("")
    return "\n".join(lines)


def send_html_email(subject, html_body, plain_body, to_email):
    """Send via Mail.app. Mirrors the weekly digest's exact MIME structure:
    set plain `content` FIRST and `html content` LAST, NO attachment. Mail.app
    produces a clean multipart/alternative that every client (Mail.app, iOS
    Mail, Outlook iPhone, Gmail) renders inline. The Excel tracker is no
    longer attached because adding it forces multipart/mixed, which causes
    Outlook to mis-render the HTML body as an .htm attachment."""
    body_tmp = tempfile.NamedTemporaryFile(
        delete=False, mode="w", suffix=".html", prefix="jobsearch_daily_"
    )
    body_tmp.write(html_body)
    body_tmp.close()

    script_path = "/tmp/jobsearch_daily_send.applescript"
    subject_escaped = subject.replace(chr(34), chr(92) + chr(34))
    plain_escaped = (
        plain_body
        .replace(chr(92), chr(92) + chr(92))
        .replace(chr(34), chr(92) + chr(34))
        .replace(chr(10), chr(92) + "n")
    )

    applescript = f'''
set bodyFile to "{body_tmp.name}"
set htmlBody to do shell script "cat " & quoted form of bodyFile
tell application "Mail"
  set newMsg to make new outgoing message with properties ¬
    {{subject:"{subject_escaped}", visible:false}}
  tell newMsg
    make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
    set content to "{plain_escaped}"
    set html content to htmlBody
  end tell
  send newMsg
end tell
'''
    with open(script_path, "w") as f:
        f.write(applescript)

    result = subprocess.run(["osascript", script_path], capture_output=True, text=True)
    os.unlink(body_tmp.name)
    if result.returncode != 0:
        raise RuntimeError(f"Email send failed: {result.stderr.strip()}")


def main():
    if len(sys.argv) < 4:
        print("Usage: send_daily_digest.py <tmp_output_path> <subject> <to_email>", file=sys.stderr)
        sys.exit(1)
    output_file, subject, to_email = sys.argv[1], sys.argv[2], sys.argv[3]

    with open(output_file, "r", errors="replace") as f:
        full_content = f.read()
    block = extract_email_block(full_content)
    summary = parse_summary(block)
    jobs = parse_jobs(block)
    today_dt = datetime.date.today()
    total_new = _safe_int(summary.get("Total new jobs", "0"))
    history_7d = collect_7day_history()
    composite_b64 = build_composite_image(today_dt, total_new, summary, history_7d)

    html_body = build_html(summary, jobs, today_dt, composite_b64)
    plain_body = build_plain_text_fallback(summary, jobs, today_dt)
    print(
        f"Parsed: {len(jobs)} jobs, summary keys: {len(summary)}, "
        f"HTML: {len(html_body):,} chars, plain: {len(plain_body):,} chars",
        file=sys.stderr,
    )
    send_html_email(subject, html_body, plain_body, to_email)
    print(f"Daily digest emailed to {to_email}")


if __name__ == "__main__":
    main()
