# Subproject: update_job Command

**Status:** Sketch / under review

---

## What it is

A new REPL command that reopens a customer's job document after the diagnostic run
and fills in the parts that weren't known at intake time:

- Section 11 of the intake form (tool/profile used, run timestamps, report filename)
- Optionally generates the Refurb Plan document for that job

---

## Proposed usage

```
sys.argv[1:] = ["update_job", "2026-023", r"C:\path\to\diagnostic_output.json"]
import runpy ; temp = runpy._run_module_as_main("repl_main")
```

Or interactively — say **"update job"** to the assistant and it will ask for the
job number and the diagnostic output file path.

---

## What it does

1. Finds `jobs/<slug>.docx` by job number (scans `jobs/` for a filename starting
   with the job number)
2. Opens the document via COM
3. Reads the diagnostic JSON (see `diagnostic_contract.md` for the format)
4. Fills Section 11 CC fields:
   - `tool_profile` ← `diagnostic.tool_profile`
   - `run_started` ← `diagnostic.run_started`
   - `run_completed` ← `diagnostic.run_completed`
   - `report_file` ← `diagnostic.report_file`
5. Saves the intake form
6. Optionally: calls `_build_refurb_plan()` to generate a companion
   `jobs/<slug>_refurb_plan.docx` pre-filled with findings and options

---

## Finding the job document

```python
def _find_job_doc(job_number: str) -> str | None:
    """Return path to jobs/<slug>.docx for the given job number, or None."""
    for fname in os.listdir(JOBS_DIR):
        if fname.startswith(job_number) and fname.endswith(".docx"):
            return os.path.join(JOBS_DIR, fname)
    return None
```

---

## Interactive update_job workflow (chat-assisted)

When the technician says "update job":

1. I ask for the job number
2. I ask for the path to the diagnostic output file
   (or the technician pastes key findings directly into chat)
3. I ask which refurb options apply (checkbox-style confirmation)
4. I ask for any custom prices
5. I run `update_job` and generate the refurb plan

---

## Open questions

- [ ] Should `update_job` also update `customers.tsv` with a "completed" timestamp?
- [ ] Should it email the refurb plan automatically, or just produce the .docx?
- [ ] What if the diagnostic JSON is missing some fields — skip silently or warn?
- [ ] Should there be a `status` field in the job (intake / diagnosed / quoted /
      completed / returned) tracked in `customers.tsv`?
