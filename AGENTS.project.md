# AGENTS.project.md — nielsoln_customers

## Project Overview

Customer management automation for the Nielsoln device repair / diagnostic service.

Automates creation of the **Device Intake Diagnostic Consent Form** (`Nielsoln_Device_Intake_Diagnostic_Consent_Form.docx`):
- Builds versioned Word templates with content controls (`build_template`)
- Fills templates with sample data for QA (`fill_sample`)
- Onboards new customers from a JSON intake file, embedding photos (`onboard_customer`)
- Maintains a race-protected job-number counter via GitHub Contents API (`next_job`)

See `docs/intake_workflow.md` for the step-by-step customer onboarding checklist.

## Development Style

**REPL-style.** See `AGENTS.base.md` (profile `python-repl-driven-dev`) for the pattern.

REPL entry point:
```
import repl_reload ; repl_reload.reload_all()
sys.argv[1:] = ["build_template"] ; import runpy ; temp = runpy._run_module_as_main("repl_main")
```

Available commands: `scan`, `build_template`, `fill_sample`, `insert_picture`, `next_job`, `onboard_customer`

See **`docs/intake_workflow.md`** for the step-by-step customer onboarding checklist.

## Repo Structure

```
nielsoln_customers/
  AGENTS.project.md       ← this file
  AGENTS.md               ← auto-generated, DO NOT EDIT
  AGENTS.project.json     ← profile declarations
  update_agents.py        ← thin shim, run to regenerate AGENTS.md
  locals.txt              ← machine-specific vars (gitignored)
  locals.txt.example      ← copy to locals.txt and fill in
  state.json              ← job number counter (committed, updated via GitHub API)
  repl_main.py            ← REPL entry point; all customer-specific business logic
  repl_reload.py          ← ordered module reload helper
  word_tools.py           ← generic Word COM automation library
  customers.tsv           ← customer database (tsv, committed)
  templates/
    Nielsoln_Device_Intake_Diagnostic_Consent_Form.docx   ← read-only source
    Nielsoln_Device_Intake_Diagnostic_Consent_Form_vNN.docx  ← built templates
  jobs/                   ← per-customer output documents (gitignored — personal data)
  docs/
    intake_workflow.md    ← customer onboarding checklist
  repl_logs/              ← per-invocation tee logs (gitignored)
```

## Module Boundaries

- **`word_tools.py`** — generic Word COM automation only. No form-specific or customer-specific
  logic. Will eventually live in its own `word_tools` repo. Import as `import word_tools`
  (not `from word_tools import`) so `importlib.reload` works correctly in the REPL.
- **`repl_main.py`** — all customer/form-specific logic: CC field definitions, checkbox layout,
  job numbering, intake workflow commands.

## Required locals.txt Variables

| Variable | Purpose |
|---|---|
| `VERSHOLN_DIR` | Path to the `versholn` sibling repo |
| `AGENT_STANDARDS_DIR` | Path to `nielsoln_agent_standards` |
| `GITHUB_TOKEN` | GitHub personal access token (Contents: read+write on this repo) |

`GITHUB_TOKEN` is optional; if absent, `next_job` falls back to local
`state.json` update (safe for single-user / offline use).
`GITHUB_REPO` is auto-detected from `git remote get-url origin` — no need to set it.

## Key Design Decisions

- Content controls use `wdContentControlRichText (Type=1)` — not `wdContentControlText (Type=2)`.
  The plain-text type blocks `Range.Text=` via COM; rich-text works correctly.
- Checkbox replacement iterates positions in **descending** order to avoid offset invalidation.
- Template versioning uses `_vNN` suffix (never modify the read-only source file).
- Job numbers use format `YYYY-NNN` (e.g. `2026-003`), tracked in `state.json`.
- Race protection for job numbers uses GitHub Contents API optimistic locking (GET + PUT with SHA;
  409 Conflict on concurrent write triggers automatic retry with back-off).
- `jobs/` is gitignored — output documents contain customer personal data.

## Notes

- `word_tools.py` will be extracted into its own repo once the API stabilises.
- `customers.tsv` is the customer database; columns: `job_number`, `date_received`,
  `customer_name`, `phone`, `email`, `address_notes`, `brand_model`, `serial_number`,
  `asset_tag`, `received_by`, `time_received`.
