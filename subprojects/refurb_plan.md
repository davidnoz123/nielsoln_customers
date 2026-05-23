# Subproject: Refurb Plan Document

**Status:** Sketch / under review

---

## What it is

A customer-facing Word document emailed (or printed) at the end of the "Free Old Laptop
Refurb Plan" appointment.  It gives the customer two things:

1. A plain-English summary of what the diagnostic found
2. A menu of refurb options with indicative pricing so they can decide what to do next

---

## Proposed document sections

### Section 1 — Device summary
Auto-filled from the intake form data (no re-entry):

| Field | Source |
|---|---|
| Customer name | intake `customer_name` |
| Device (brand/model) | intake `brand_model` |
| Serial number | intake `serial_number` |
| Date assessed | intake `date_received` |
| Job number | intake `job_number` |
| Assessed by | intake `received_by` |

### Section 2 — Diagnostic findings
Populated from the diagnostic tool output (see `diagnostic_contract.md`):

| Finding | Value | Status |
|---|---|---|
| Storage health | e.g. "SMART: OK — 94% life remaining" | ✅ / ⚠️ / ❌ |
| RAM | e.g. "8 GB — no errors detected" | ✅ / ⚠️ / ❌ |
| Battery (laptops) | e.g. "42% wear — 3.1 Wh remaining of 5.4 Wh design" | ✅ / ⚠️ / ❌ |
| OS integrity | e.g. "Windows 10 — minor file corruption detected" | ✅ / ⚠️ / ❌ |
| Malware scan | e.g. "ClamAV quick scan: clean" | ✅ / ⚠️ / ❌ |
| Thermals | e.g. "CPU 71°C under load — within spec" | ✅ / ⚠️ / ❌ |

Free-text "Technician notes" field for anything not covered above.

### Section 3 — Refurb options
Each option is a checkbox row. Technician ticks the ones relevant for this device.
Pricing is a CC field (editable per job):

| Option | Description | Indicative price |
|---|---|---|
| `opt_ssd_clone` | Clone existing HDD → new SSD (same or larger capacity) | $[price] |
| `opt_ssd_upgrade` | Replace HDD/SSD with larger capacity SSD | $[price] |
| `opt_ram_upgrade` | Upgrade RAM to [capacity] GB | $[price] |
| `opt_os_reinstall` | Clean Windows reinstall (data backed up first) | $[price] |
| `opt_os_upgrade` | Upgrade to Windows 11 | $[price] |
| `opt_battery_replace` | Replace battery | $[price] |
| `opt_screen_replace` | Replace screen | $[price] |
| `opt_thermal_paste` | Replace thermal paste / clean fans | $[price] |
| `opt_data_backup` | Full data backup to external drive | $[price] |
| `opt_keep_as_is` | No action recommended — device is healthy | — |

### Section 4 — Next steps
Short paragraph or bullet list. Could be a free-text CC or boilerplate based on
which options were selected.

### Section 5 — Technician sign-off
Technician name, date, contact details.

---

## Template file

`templates/Nielsoln_Refurb_Plan.docx` — source (read-only)
`templates/Nielsoln_Refurb_Plan_vNN.docx` — built version with CCs

---

## Open questions

- [ ] Does the customer receive a printed copy, an emailed PDF, or both?
- [ ] Are prices fixed/standard or do they vary per job?
- [ ] Should the findings traffic-light (✅ ⚠️ ❌) be checkboxes in Word or just text?
- [ ] Do we want a "total estimate" field summing the selected options?
- [ ] Should the refurb plan reference the intake job number so both docs are linked?
