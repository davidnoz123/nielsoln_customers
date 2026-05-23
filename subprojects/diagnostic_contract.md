# Subproject: Diagnostic Output Contract

**Status:** Sketch / under review

---

## What it is

A minimal JSON format that `nielsoln_rescue_toolkit` and `usb_device_tools` write
at the end of a diagnostic run so that `nielsoln_customers` can consume it without
importing or knowing anything about those repos.

`nielsoln_customers` is the consumer.  The diagnostic repos are the producers.
No code dependency in either direction — only this JSON format.

---

## Proposed format

```json
{
  "schema_version": 1,
  "job_number":     "2026-023",
  "tool_profile":   "nielsoln_rescue_toolkit quick-scan v0.1.14",
  "run_started":    "23/05/2026 14:35",
  "run_completed":  "23/05/2026 14:52",
  "report_file":    "2026-023_quick-scan_20260523_1452.txt",

  "findings": {
    "storage_health":  { "status": "ok",      "detail": "SMART: 94% life remaining" },
    "ram":             { "status": "ok",       "detail": "8 GB — no errors" },
    "battery":         { "status": "warning",  "detail": "42% wear — 3.1 Wh of 5.4 Wh" },
    "os_integrity":    { "status": "warning",  "detail": "Minor file corruption detected" },
    "malware_scan":    { "status": "ok",       "detail": "ClamAV quick scan: clean" },
    "thermals":        { "status": "ok",       "detail": "CPU 71°C under load — within spec" }
  },

  "device_identity": {
    "serial_number":   "SN12345678",
    "make_model":      "Dell XPS 15 9570",
    "os_version":      "Windows 10 22H2"
  },

  "technician_notes": ""
}
```

### Status values

| Value | Meaning |
|---|---|
| `"ok"` | No action required |
| `"warning"` | Monitor or consider action |
| `"critical"` | Action recommended / device at risk |
| `"not_run"` | This check was skipped or not applicable |

---

## Output file naming convention

```
<job_number>_diagnostic_<YYYYMMDD_HHMM>.json
```

Example: `2026-023_diagnostic_20260523_1452.json`

The job number must be entered by the technician when starting the scan so it
is embedded in the output.  This is the only coupling between repos.

---

## Where files are written

Each diagnostic repo writes to its own output area:

- **`nielsoln_rescue_toolkit`:** USB drive `logs/` directory
- **`usb_device_tools`:** `evidence\<timestamp>\` directory

The technician copies or points `update_job` at the file — no automatic sync needed.

---

## Cross-repo reference

This contract should be referenced (not duplicated) in:
- `nielsoln_rescue_toolkit/AGENTS.project.md` — "emit this JSON at run end"
- `usb_device_tools/AGENTS.project.md` — "emit this JSON at run end"
- `nielsoln_customers/docs/intake_workflow.md` — "update_job reads this format"

---

## Open questions

- [ ] Should `findings` be a fixed set of keys (as above) or a free-form list?
      Fixed keys make the refurb plan template easier to design.
- [ ] Does `usb_device_tools` have different finding categories (e.g. no `battery`
      for a desktop)? Handle with `"status": "not_run"` or omit the key?
- [ ] Should the schema version bump require a `nielsoln_customers` code change,
      or should the consumer be tolerant of unknown keys?
- [ ] Do we want a `"recommended_options"` list pre-populated by the toolkit
      (e.g. `["opt_ssd_clone", "opt_thermal_paste"]`) for the technician to review?
