"""Unit tests for the pure parser functions in send_daily_digest.py.

These functions handle the LLM agent's structured-block output and the URL
normalization that protects the digest's "View posting" buttons from broken
links. Anything that fails here would produce a broken or misleading email.
"""
import os
import sys
import unittest

# send_daily_digest reads BASE_DIR at import time via _require_env; provide a
# placeholder so the import succeeds in CI without a populated .env.
os.environ.setdefault("BASE_DIR", "/tmp/jobsearch-tests")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import send_daily_digest as sdd  # noqa: E402


class ExtractEmailBlockTests(unittest.TestCase):
    def test_extracts_between_markers(self):
        content = "preamble\n<<<EMAIL_START>>>\nbody line\n<<<EMAIL_END>>>\npostscript"
        self.assertEqual(sdd.extract_email_block(content), "body line")

    def test_start_marker_only_returns_tail(self):
        content = "junk\n<<<EMAIL_START>>>\nthe rest of the file"
        self.assertEqual(sdd.extract_email_block(content), "the rest of the file")

    def test_no_markers_returns_whole_content_stripped(self):
        self.assertEqual(sdd.extract_email_block("  hello  "), "hello")


class ParseSummaryTests(unittest.TestCase):
    def test_parses_known_keys(self):
        block = (
            "Spain jobs found: 3\n"
            "US Remote jobs found: 2\n"
            "LATAM jobs found: 1\n"
            "Total new jobs: 6\n"
            "Folders created: 6\n"
            "Jobs skipped (dup): 4\n"
            "Jobs skipped (old): 2\n"
            "Jobs skipped (filter): 1\n"
        )
        out = sdd.parse_summary(block)
        self.assertEqual(out["Spain jobs found"], "3")
        self.assertEqual(out["Total new jobs"], "6")
        self.assertEqual(out["Jobs skipped (filter)"], "1")

    def test_missing_keys_omitted(self):
        out = sdd.parse_summary("Spain jobs found: 1\n")
        self.assertIn("Spain jobs found", out)
        self.assertNotIn("Total new jobs", out)


class ParseJobsTests(unittest.TestCase):
    def test_parses_one_job(self):
        block = (
            "NEW JOBS DETAIL\n\n"
            "[SCORE: 5/5] Senior Power Automate Developer @ Telefónica\n"
            "Market: Spain\n"
            "Location: Madrid, Spain\n"
            "Mode: Hybrid\n"
            "Salary: €45,000–€55,000\n"
            "Posted: 2 days ago\n"
            "Board: LinkedIn\n"
            "URL: https://www.linkedin.com/jobs/view/3849273948/\n"
            "Why it matches: Strong Power Automate alignment with hybrid Madrid base.\n"
        )
        jobs = sdd.parse_jobs(block)
        self.assertEqual(len(jobs), 1)
        j = jobs[0]
        self.assertEqual(j["score"], 5)
        self.assertEqual(j["title"], "Senior Power Automate Developer")
        self.assertEqual(j["company"], "Telefónica")
        self.assertEqual(j["market"], "Spain")
        self.assertEqual(j["url"], "https://www.linkedin.com/jobs/view/3849273948/")
        self.assertTrue(j["why"].startswith("Strong Power Automate"))

    def test_parses_multiple_jobs(self):
        block = (
            "NEW JOBS DETAIL\n\n"
            "[SCORE: 5/5] Title A @ Company A\nURL: https://a.example/1\n\n"
            "[SCORE: 3/5] Title B @ Company B\nURL: https://b.example/2\n"
        )
        jobs = sdd.parse_jobs(block)
        self.assertEqual(len(jobs), 2)
        self.assertEqual(jobs[0]["score"], 5)
        self.assertEqual(jobs[1]["score"], 3)

    def test_no_detail_section_returns_empty(self):
        self.assertEqual(sdd.parse_jobs("Spain jobs found: 0\n"), [])

    def test_malformed_score_line_skipped(self):
        block = "NEW JOBS DETAIL\n\n[SCORE: garbage] Title @ Company\n"
        self.assertEqual(sdd.parse_jobs(block), [])


class ValidateAndNormalizeUrlTests(unittest.TestCase):
    def test_none_and_empty_rejected(self):
        self.assertIsNone(sdd.validate_and_normalize_url(None))
        self.assertIsNone(sdd.validate_and_normalize_url(""))
        self.assertIsNone(sdd.validate_and_normalize_url("   "))

    def test_placeholder_tokens_rejected(self):
        for tok in ("n/a", "N/A", "none", "TBD", "tba", "unknown"):
            self.assertIsNone(sdd.validate_and_normalize_url(tok), tok)

    def test_linkedin_canonicalized(self):
        # All variants collapse to the canonical /jobs/view/{id}/ form.
        cases = [
            "https://www.linkedin.com/jobs/view/3849273948/",
            "https://linkedin.com/jobs/view/3849273948",
            "https://www.linkedin.com/comm/jobs/view/3849273948?refId=xyz",
            "https://es.linkedin.com/jobs/view/3849273948?trk=foo",
        ]
        for raw in cases:
            self.assertEqual(
                sdd.validate_and_normalize_url(raw),
                "https://www.linkedin.com/jobs/view/3849273948/",
                raw,
            )

    def test_linkedin_id_too_short_not_canonicalized(self):
        # 8-12 digits are required to match the LinkedIn canonicalizer. A
        # too-short ID (truncation symptom — see extract_linkedin_urls.py for
        # the quoted-printable explanation) falls through to the generic-URL
        # path: not canonicalized, but also not rejected here.
        raw = "https://www.linkedin.com/jobs/view/4414/"
        self.assertEqual(sdd.validate_and_normalize_url(raw), raw)

    def test_search_and_list_urls_rejected(self):
        for raw in (
            "https://www.linkedin.com/jobs/search?keywords=python",
            "https://www.linkedin.com/jobs/collections/recommended",
            "https://www.google.com/search?q=python+developer",
            "https://www.indeed.com/jobs?q=python",
            "https://www.infojobs.net/ofertas-trabajo/madrid",
        ):
            self.assertIsNone(sdd.validate_and_normalize_url(raw), raw)

    def test_generic_url_passes_through(self):
        self.assertEqual(
            sdd.validate_and_normalize_url("https://tecnoempleo.com/empleo/12345"),
            "https://tecnoempleo.com/empleo/12345",
        )

    def test_trailing_punctuation_stripped(self):
        self.assertEqual(
            sdd.validate_and_normalize_url("https://example.com/job."),
            "https://example.com/job",
        )
        self.assertEqual(
            sdd.validate_and_normalize_url("(https://example.com/job)"),
            "https://example.com/job",
        )


if __name__ == "__main__":
    unittest.main()
