# Customer Intake Workflow

This document describes the step-by-step checklist for onboarding a new customer.
It is designed to be followed interactively with the AI assistant in the VS Code chat.

---

## How to start

Tell the assistant: **"new customer"**

The assistant will work through the steps below, ask for each piece of information,
accept photos where noted, and produce a completed `jobs/<job_number>_<name>.docx`
at the end.

---

## Step 1 — Job setup

| Field | Notes |
|---|---|
| Job number | Claimed automatically via `next_job` (race-protected; see below) |
| Date received | Default: today |
| Received by | Your name or initials |
| Time received | Default: now (hh:mm) |

## Step 2 — Customer details

| Field | Notes |
|---|---|
| Full name | Used in form and filename slug |
| Phone | |
| Email | |
| Preferred contact | Phone / Text / Email |
| Urgency | Normal / Urgent |
| Address / notes | Optional |

## Step 3 — Device

| Field | Notes |
|---|---|
| Device type | Laptop / Desktop / Phone / Tablet / Other |
| Brand and model | **Photo accepted** — paste a photo of the label |
| Serial number | **Photo accepted** — paste a photo of the serial number sticker |
| Asset tag / device name | Optional |

## Step 4 — Condition and accessories

| Field | Notes |
|---|---|
| Operating system | Windows / macOS / Linux / iOS / Android |
| Power status | Starts / No power / Intermittent |
| Password / PIN supplied | Yes / No |
| Charger received | Yes / No |
| Accessories | Bag / Mouse / Keyboard / USB drive / Other |
| External damage | Yes / No — **damage photo accepted** |
| Device front photo | **Photo accepted** |

## Step 5 — Customer acknowledgement

| Field | Notes |
|---|---|
| Customer signature | **Photo of signed paper accepted** |
| Printed name | Auto-filled from customer name |
| Date | Auto-filled from today |

## Step 6 — Generate document

The assistant compiles a JSON intake file and runs:

```
sys.argv[1:] = ["onboard_customer", "<path/to/intake.json>"]
import runpy ; temp = runpy._run_module_as_main("repl_main")
```

Output: `jobs/<job_number>_<customer_name>.docx`

---

## Photo workflow

When you paste a photo into the chat:

1. The assistant reads text from it (brand, model, serial number, damage description).
2. You save the image to a local path (e.g. `C:\intake_photos\img001.jpg`).
3. The path is included in the JSON file so `onboard_customer` can embed it.

Supported picture fields: `sig_customer`, `sig_customer_return`, `sig_technician`,
`device_front`, `damage_photo`.

---

## Job number tracking

Job numbers use the format **`YYYY-NNN`** (e.g. `2026-003`).
The counter lives in `state.json` at the repo root, which is committed to git.

To claim the next job number:

```
sys.argv[1:] = ["next_job"]
import runpy ; temp = runpy._run_module_as_main("repl_main")
```

### Race protection via GitHub API

If `GITHUB_TOKEN` and `GITHUB_REPO` are set in `locals.txt`, the counter is
updated through the **GitHub Contents API** rather than editing the local file
directly. This provides optimistic locking:

1. `GET /repos/{owner}/{repo}/contents/state.json` → receive current content + SHA
2. Increment `last_job_number` in memory
3. `PUT /repos/{owner}/{repo}/contents/state.json` with the current SHA
4. If another client committed between the GET and PUT, GitHub returns **409 Conflict** —
   the agent retries automatically (up to 5 times with back-off)

This means two people working simultaneously can never claim the same job number,
as long as both have network access to the GitHub repo.

If `GITHUB_TOKEN`/`GITHUB_REPO` are absent, the code falls back to updating
`state.json` locally. This is safe for a single-person workflow or offline use.

### Setting up GitHub access

Add to `locals.txt` (this file is gitignored — never commit it):

```
GITHUB_TOKEN=ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
GITHUB_REPO=your-github-username/nielsoln_customers
```

Generate a token at: GitHub → Settings → Developer settings →
Personal access tokens → Fine-grained tokens.
Required permission: **Contents** → Read and write (on this repo only).

---

## Intake JSON format

`onboard_customer` accepts a flat JSON object with any combination of these keys:

### Text fields

```json
{
  "job_number":    "2026-001",
  "date_received": "23/05/2026",
  "customer_name": "Jane Smith",
  "phone":         "0400 123 456",
  "email":         "jane@example.com",
  "address_notes": "42 Example St",
  "brand_model":   "Dell XPS 15",
  "serial_number": "SN12345678",
  "asset_tag":     "JANE-LAPTOP",
  "received_by":   "Tech 1",
  "time_received": "14:30",
  "printed_name":  "Jane Smith",
  "sig_date":      "23/05/2026",
  "date_returned": "",
  "tool_profile":  "",
  "run_started":   "",
  "run_completed": "",
  "report_file":   ""
}
```

### Checkbox fields (true/false)

```json
{
  "pref_contact_phone": false,
  "pref_contact_text":  false,
  "pref_contact_email": true,
  "urgency_normal":     true,
  "urgency_urgent":     false,
  "device_laptop":      true,
  "os_windows":         true,
  "power_starts":       true,
  "password_no":        true,
  "charger_yes":        true,
  "damage_no":          true
}
```

### Picture fields (absolute image paths)

```json
{
  "sig_customer":  "C:\\intake_photos\\sig_jane.jpg",
  "device_front":  "C:\\intake_photos\\device_front.jpg",
  "damage_photo":  "C:\\intake_photos\\damage.jpg"
}
```
