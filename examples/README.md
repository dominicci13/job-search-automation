# Examples

This directory exists as a placeholder for files you bring yourself:

- `resume_base.docx` — your clean ATS-friendly resume. The agent tailors
  this per role; place it at `$BASE_DIR/resume_base.docx`.
- `cv_base_es.docx` — (optional) Spanish version for Spain-market jobs.

Both are gitignored by default to keep your personal data out of the repo.
If you want to commit example resumes for portfolio purposes, place them
under `examples/` (this directory is allowlisted in `.gitignore`).

## Excel tracker

Don't ship a populated tracker. Run `python scripts/setup_excel.py` to create
a blank one at `$BASE_DIR/job_tracker.xlsx` after cloning.
