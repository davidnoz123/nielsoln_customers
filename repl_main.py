"""repl_main.py — Customer management automation workflows.

REPL usage:
    import repl_reload ; repl_reload.reload_all()
    sys.argv[1:] = ["scan"] ; import runpy ; temp = runpy._run_module_as_main("repl_main")
    sys.argv[1:] = ["scan", "path/to/doc.docx"] ; import runpy ; temp = runpy._run_module_as_main("repl_main")

Commands:
    scan [path]           -- Scan a Word document and print its structure.
                             Defaults to templates/Nielsoln_Device_Intake_Diagnostic_Consent_Form.docx
    build_template        -- Build a versioned template with content controls (_vNN.docx).
    fill_sample           -- Populate the latest template version with sample data (_vNN_sample.docx).
    onboard_customer <json_path>
                          -- Fill a new intake form from a JSON data file.
                             Produces jobs/<job_number>_<customer_slug>.docx
                             JSON keys: see _INTAKE_FIELDS below for all supported tags.
                             Picture tags (sig_customer, device_front, damage_photo, etc.) should
                             map to absolute image file paths.
    insert_picture <tag> <image_path>
                          -- Embed an image into a picture CC by tag in the latest sample document.
                             Example tags: sig_customer, sig_customer_return, sig_technician,
                                           device_front, damage_photo
    capture_sig <tag> <job_doc_path>
                          -- Open a signature-capture window (Wacom tablet or mouse).
                             On Accept, embeds the PNG directly into the named CC in the job doc.
                             Example: capture_sig sig_customer jobs/2026-023_jane-smith.docx

Customer onboarding checklist: docs/intake_workflow.md
"""

import contextlib
import datetime
import os
import sys

_log = lambda msg: print(msg, flush=True)


def _install_and_import(package: str, pip_name: str = None):
    """Import *package*, installing via pip if absent. Returns the module."""
    import importlib as _il
    import subprocess as _sp
    try:
        return _il.import_module(package)
    except ImportError:
        _log(f"installing {pip_name or package} ...")
        _sp.check_call(
            [sys.executable, "-m", "pip", "install", pip_name or package],
            stdout=_sp.DEVNULL,
        )
        return _il.import_module(package)


def _inject_picture_into_cc(doc_path: str, tag: str, image_path: str) -> bool:
    """Insert an image into a picture CC by direct docx XML manipulation.

    Bypasses Word COM's AddPicture guard.  The document must NOT be open
    in Word when this is called (Windows holds a write lock on open docs).
    Returns True if the tag was found and the image was injected.
    """
    import copy
    docx_mod = _install_and_import("docx", "python-docx")
    from docx.oxml.ns import qn as _qn
    from lxml import etree as _et

    d = docx_mod.Document(doc_path)

    # Add the image via a temporary paragraph so python-docx handles
    # all relationship / image-part plumbing for us.
    tmp_para = d.add_paragraph()
    tmp_run  = tmp_para.add_run()
    tmp_run.add_picture(image_path)
    drawing_elem = tmp_run._element.find(_qn("w:drawing"))
    if drawing_elem is None:
        d.element.body.remove(tmp_para._element)
        return False
    drawing_copy = copy.deepcopy(drawing_elem)
    d.element.body.remove(tmp_para._element)

    # Find the SDT by tag (searches whole body tree inc. table cells).
    target_sdt = None
    for sdt in d.element.body.iter(_qn("w:sdt")):
        sdt_pr = sdt.find(_qn("w:sdtPr"))
        if sdt_pr is not None:
            tag_elem = sdt_pr.find(_qn("w:tag"))
            if tag_elem is not None and tag_elem.get(_qn("w:val")) == tag:
                target_sdt = sdt
                break
    if target_sdt is None:
        return False

    # Remove 'showingPlcHdr' flag so Word shows content, not placeholder.
    sdt_pr = target_sdt.find(_qn("w:sdtPr"))
    if sdt_pr is not None:
        plc = sdt_pr.find(_qn("w:showingPlcHdr"))
        if plc is not None:
            sdt_pr.remove(plc)

    # Replace sdtContent with a paragraph containing the drawing.
    old_content = target_sdt.find(_qn("w:sdtContent"))
    if old_content is not None:
        target_sdt.remove(old_content)
    sdt_content = _et.SubElement(target_sdt, _qn("w:sdtContent"))
    p_elem      = _et.SubElement(sdt_content, _qn("w:p"))
    r_elem      = _et.SubElement(p_elem, _qn("w:r"))
    r_elem.append(drawing_copy)

    d.save(doc_path)
    return True


REPO_ROOT     = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(REPO_ROOT, "templates")
JOBS_DIR      = os.path.join(REPO_ROOT, "jobs")
STATE_JSON    = os.path.join(REPO_ROOT, "state.json")
CUSTOMERS_TSV = os.path.join(REPO_ROOT, "customers.tsv")
JOBS_TSV      = os.path.join(REPO_ROOT, "jobs.tsv")
DEFAULT_TEMPLATE = os.path.join(
    TEMPLATES_DIR, "Nielsoln_Device_Intake_Diagnostic_Consent_Form.docx"
)


# ---------------------------------------------------------------------------
# Template field definitions  (customer-specific business logic)
# ---------------------------------------------------------------------------

# Text content controls to insert into blank table cells.
# Each entry: (table_index, row, col, tag, title, placeholder)
# All indices 1-based; table numbering matches the source document scan.
_TEXT_CC_FIELDS = [
    # ─ Section 1: Customer details (Table 2, 4 rows × 4 cols) ─────────────────────
    (2, 1, 2, "customer_name",  "Customer Name",   "Customer name"),
    (2, 1, 4, "date_received",  "Date Received",   "dd/mm/yyyy"),
    (2, 2, 2, "phone",          "Phone",           "Phone number"),
    (2, 2, 4, "email",          "Email",           "Email address"),
    (2, 4, 2, "address_notes",  "Address / Notes", "Address or notes"),
    (2, 4, 4, "job_number",     "Job Number",      "Job number"),
    # ─ Section 2: Device details (Table 3, 5 rows × 4 cols) ─────────────────────
    (3, 1, 4, "brand_model",    "Brand / Model",   "Make and model"),
    (3, 2, 2, "serial_number",  "Serial Number",   "Serial number"),
    (3, 2, 4, "asset_tag",      "Asset Tag / Name", "Asset tag or device name"),
    # ─ Section 10: Technician intake (Table 9, 3 rows × 4 cols) ────────────────
    (9, 1, 2, "received_by",    "Received By",     "Technician name"),
    (9, 1, 4, "time_received",  "Time Received",   "hh:mm"),
    # ─ Section 11: Automated health-check record (Table 10, 5 rows × 4 cols) ──
    (10, 1, 2, "tool_profile",  "Tool / Profile Used", "Tool or profile name"),
    (10, 1, 4, "run_started",   "Run Started",         "dd/mm/yyyy hh:mm"),
    (10, 2, 2, "run_completed", "Run Completed",       "dd/mm/yyyy hh:mm"),
    (10, 2, 4, "report_file",   "Report File / Name",  "Report filename"),
]

# Paragraph-level text CCs: (prefix_text, tag, title)
# The prefix_text must exactly match the start of the target paragraph.
# Trailing content (underscores etc.) is deleted and a hidden text CC is inserted.
_PARA_TEXT_CC_FIELDS = [
    # ─ Section 9: Customer acknowledgement ─────────────────────────────────────
    ("Printed name: ",  "printed_name", "Printed Name"),
    ("Date: ",          "sig_date",     "Signature Date"),
    # ─ Section 12: Device return ────────────────────────────────────────────────
    ("Date returned: ", "date_returned", "Date Returned"),
]

# Paragraph-level picture CCs: (prefix_text, tag, title)
# The prefix_text must exactly match the start of the target paragraph.
# Trailing content is deleted and a picture CC is inserted after the prefix.
_PARA_PICTURE_CC_FIELDS = [
    # ─ Section 9: Customer signature ────────────────────────────────────────────
    ("Customer signature: ",         "sig_customer",        "Customer Signature"),
    # ─ Section 12: Return signatures ────────────────────────────────────────────
    ("Customer received signature: ", "sig_customer_return", "Customer Return Signature"),
    ("Technician signature: ",        "sig_technician",      "Technician Signature"),
]

# Paragraph-level picture CCs for device photos: (prefix_text, tag, title)
# These are NEW paragraphs appended by build_template (no existing paragraph to find).
# Handled separately in _cmd_build_template after the form body.
_DEVICE_PHOTO_TAGS = [
    ("device_front",  "Device Front Photo"),
    ("damage_photo",  "Device Damage Photo"),
]

# Checkbox content controls to replace \u2610 chars in specific table cells.
# Each entry: (table_index, row, col, [(tag, title), ...])
# The (tag, title) list must match the left-to-right order of \u2610 chars in the cell.
_CELL_CHECKBOXES = [
    # ─ Section 1: Preferred contact ────────────────────────────────────────────
    (2, 3, 2, [
        ("pref_contact_phone", "Preferred Contact: Phone"),
        ("pref_contact_text",  "Preferred Contact: Text"),
        ("pref_contact_email", "Preferred Contact: Email"),
    ]),
    # ─ Section 1: Urgency ─────────────────────────────────────────────────────
    (2, 3, 4, [
        ("urgency_normal", "Urgency: Normal"),
        ("urgency_urgent", "Urgency: Urgent"),
    ]),
    # ─ Section 2: Device type ───────────────────────────────────────────────
    (3, 1, 2, [
        ("device_laptop",  "Device: Laptop"),
        ("device_desktop", "Device: Desktop"),
        ("device_phone",   "Device: Phone"),
        ("device_tablet",  "Device: Tablet"),
        ("device_other",   "Device: Other"),
    ]),
    # ─ Section 2: Operating system ──────────────────────────────────────────
    (3, 3, 2, [
        ("os_windows", "OS: Windows"),
        ("os_macos",   "OS: macOS"),
        ("os_linux",   "OS: Linux"),
        ("os_ios",     "OS: iOS"),
        ("os_android", "OS: Android"),
    ]),
    # ─ Section 2: Power status ──────────────────────────────────────────────
    (3, 3, 4, [
        ("power_starts",       "Power: Starts"),
        ("power_no_power",     "Power: No power"),
        ("power_intermittent", "Power: Intermittent"),
    ]),
    # ─ Section 2: Password/PIN supplied ─────────────────────────────────────
    (3, 4, 2, [
        ("password_no",  "Password: No"),
        ("password_yes", "Password: Yes"),
    ]),
    # ─ Section 2: Charger received ─────────────────────────────────────────
    (3, 4, 4, [
        ("charger_no",  "Charger: No"),
        ("charger_yes", "Charger: Yes"),
    ]),
    # ─ Section 2: Accessories received ─────────────────────────────────────
    (3, 5, 2, [
        ("acc_bag",      "Accessories: Bag"),
        ("acc_mouse",    "Accessories: Mouse"),
        ("acc_keyboard", "Accessories: Keyboard"),
        ("acc_usb",      "Accessories: USB drive"),
        ("acc_other",    "Accessories: Other"),
    ]),
    # ─ Section 2: External damage ───────────────────────────────────────────
    (3, 5, 4, [
        ("damage_no",  "External Damage: No"),
        ("damage_yes", "External Damage: Yes"),
    ]),
    # ─ Section 10: Technician risk levels ───────────────────────────────────
    (9, 2, 2, [
        ("risk_level_low",    "Initial Risk: Low"),
        ("risk_level_medium", "Initial Risk: Medium"),
        ("risk_level_high",   "Initial Risk: High"),
    ]),
    (9, 2, 4, [
        ("data_risk_low",    "Data Risk: Low"),
        ("data_risk_medium", "Data Risk: Medium"),
        ("data_risk_high",   "Data Risk: High"),
    ]),
    # ─ Section 10: Customer copy given ──────────────────────────────────────
    (9, 3, 4, [
        ("copy_given_yes", "Customer Copy: Yes"),
        ("copy_given_no",  "Customer Copy: No"),
    ]),
]


# ---------------------------------------------------------------------------
# Locals and job numbering
# ---------------------------------------------------------------------------

def _load_locals() -> dict:
    """Parse locals.txt (KEY=VALUE lines) and return as dict. Non-fatal if missing."""
    path = os.path.join(REPO_ROOT, "locals.txt")
    result = {}
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    result[key.strip()] = value.strip()
    except FileNotFoundError:
        pass
    return result


def _github_claim_next_job(github_repo: str, github_token: str, max_retries: int = 5) -> str:
    """Atomically claim the next job number via GitHub Contents API.

    Uses optimistic locking: GET state.json (receives current content + SHA),
    increment counter in memory, PUT back with the SHA.  If another client
    committed between the GET and PUT, GitHub returns 409 Conflict — the
    function retries automatically with exponential back-off.

    Returns the new job ID string, e.g. '2026-003'.
    Raises RuntimeError if max_retries exhausted.
    """
    import base64
    import json
    import time
    import urllib.error
    import urllib.request

    url = f"https://api.github.com/repos/{github_repo}/contents/state.json"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "nielsoln_customers",
    }

    for attempt in range(1, max_retries + 1):
        # ── GET current state ────────────────────────────────────────────────
        req = urllib.request.Request(url, headers=headers)
        current_sha = None
        state = None
        try:
            with urllib.request.urlopen(req) as resp:
                file_data = json.loads(resp.read())
            current_sha = file_data["sha"]
            state = json.loads(base64.b64decode(file_data["content"]))
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                # File doesn't exist on GitHub yet — seed it
                state = {"year": 0, "last_job_number": 0}
            else:
                raise

        # ── Increment ────────────────────────────────────────────────────────
        today_year = datetime.datetime.now().year
        if state.get("year") != today_year:
            state["year"] = today_year
            state["last_job_number"] = 0
        state["last_job_number"] += 1
        job_id = f"{state['year']}-{state['last_job_number']:03d}"

        # ── PUT back ─────────────────────────────────────────────────────────
        put_body = {
            "message": f"claim job {job_id}",
            "content": base64.b64encode(
                json.dumps(state, separators=(",", ":")).encode()
            ).decode(),
        }
        if current_sha is not None:
            put_body["sha"] = current_sha
        payload = json.dumps(put_body).encode()
        put_req = urllib.request.Request(
            url, data=payload, headers={**headers, "Content-Type": "application/json"},
            method="PUT",
        )
        try:
            with urllib.request.urlopen(put_req):
                pass
            _log(f"  [job claim] {job_id} claimed via GitHub (attempt {attempt})")
            return job_id
        except urllib.error.HTTPError as exc:
            if exc.code == 409 and attempt < max_retries:
                _log(f"  [job claim] conflict on attempt {attempt}, retrying...")
                time.sleep(0.5 * attempt)
                continue
            raise

    raise RuntimeError(f"Failed to claim job number after {max_retries} attempts")


def _detect_github_repo() -> str:
    """Derive 'owner/repo' from the git remote URL.  Returns '' on failure."""
    import re
    import subprocess
    try:
        url = subprocess.check_output(
            ["git", "-C", REPO_ROOT, "remote", "get-url", "origin"],
            stderr=subprocess.DEVNULL,
        ).decode().strip()
        # https://github.com/owner/repo.git  or  git@github.com:owner/repo.git
        m = re.search(r"github\.com[:/](.+?)(?:\.git)?$", url)
        return m.group(1) if m else ""
    except Exception:
        return ""


def _next_job_number() -> str:
    """Claim the next job number.

    If GITHUB_TOKEN is set in locals.txt, the claim is race-protected via the
    GitHub Contents API (optimistic locking with SHA).  The repo is detected
    automatically from the git remote — no need to set GITHUB_REPO.
    Falls back to a local state.json increment if token is absent.

    Returns the new job ID string, e.g. '2026-003'.
    """
    import json

    cfg   = _load_locals()
    token = cfg.get("GITHUB_TOKEN", "").strip()
    repo  = cfg.get("GITHUB_REPO", "").strip() or _detect_github_repo()

    if token and repo and not token.startswith("ghp_xxx"):
        return _github_claim_next_job(repo, token)

    # ── Local fallback ───────────────────────────────────────────────────────
    _log("  [job claim] GITHUB_TOKEN/GITHUB_REPO not configured — using local state.json")
    try:
        with open(STATE_JSON, encoding="utf-8") as fh:
            state = json.load(fh)
    except FileNotFoundError:
        state = {"year": 0, "last_job_number": 0}

    today_year = datetime.datetime.now().year
    if state.get("year") != today_year:
        state["year"] = today_year
        state["last_job_number"] = 0
    state["last_job_number"] += 1
    job_id = f"{state['year']}-{state['last_job_number']:03d}"

    with open(STATE_JSON, "w", encoding="utf-8") as fh:
        json.dump(state, fh)
    _log(f"  [job claim] {job_id} claimed locally")
    return job_id


def _github_claim_next_customer(github_repo: str, github_token: str, max_retries: int = 5) -> str:
    """Atomically claim the next customer ID via GitHub Contents API.

    Same optimistic-locking approach as _github_claim_next_job.
    Returns e.g. 'C-001'.
    """
    import base64
    import json
    import time
    import urllib.error
    import urllib.request

    url = f"https://api.github.com/repos/{github_repo}/contents/state.json"
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "nielsoln_customers",
    }

    for attempt in range(1, max_retries + 1):
        req = urllib.request.Request(url, headers=headers)
        current_sha = None
        state = None
        try:
            with urllib.request.urlopen(req) as resp:
                file_data = json.loads(resp.read())
            current_sha = file_data["sha"]
            state = json.loads(base64.b64decode(file_data["content"]))
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                state = {}
            else:
                raise

        state.setdefault("last_customer_id", 0)
        state["last_customer_id"] += 1
        cid = f"C-{state['last_customer_id']:03d}"

        put_body = {
            "message": f"claim customer {cid}",
            "content": base64.b64encode(
                json.dumps(state, separators=(",", ":")).encode()
            ).decode(),
        }
        if current_sha is not None:
            put_body["sha"] = current_sha
        payload = json.dumps(put_body).encode()
        put_req = urllib.request.Request(
            url, data=payload, headers={**headers, "Content-Type": "application/json"},
            method="PUT",
        )
        try:
            with urllib.request.urlopen(put_req):
                pass
            _log(f"  [customer claim] {cid} claimed via GitHub (attempt {attempt})")
            return cid
        except urllib.error.HTTPError as exc:
            if exc.code == 409 and attempt < max_retries:
                _log(f"  [customer claim] conflict on attempt {attempt}, retrying...")
                time.sleep(0.5 * attempt)
                continue
            raise

    raise RuntimeError(f"Failed to claim customer ID after {max_retries} attempts")


def _next_customer_id() -> str:
    """Claim the next customer ID.

    Uses GitHub Contents API if GITHUB_TOKEN is configured, otherwise local
    state.json fallback.
    Returns e.g. 'C-001'.
    """
    import json

    cfg   = _load_locals()
    token = cfg.get("GITHUB_TOKEN", "").strip()
    repo  = cfg.get("GITHUB_REPO", "").strip() or _detect_github_repo()

    if token and repo and not token.startswith("ghp_xxx"):
        return _github_claim_next_customer(repo, token)

    _log("  [customer claim] GITHUB_TOKEN/GITHUB_REPO not configured — using local state.json")
    try:
        with open(STATE_JSON, encoding="utf-8") as fh:
            state = json.load(fh)
    except FileNotFoundError:
        state = {}

    state.setdefault("last_customer_id", 0)
    state["last_customer_id"] += 1
    cid = f"C-{state['last_customer_id']:03d}"

    with open(STATE_JSON, "w", encoding="utf-8") as fh:
        json.dump(state, fh)
    _log(f"  [customer claim] {cid} claimed locally")
    return cid


# ---------------------------------------------------------------------------
# Log capture (tee stdout/stderr to a timestamped file per invocation)
# ---------------------------------------------------------------------------

class _TeeStream:
    def __init__(self, primary, secondary):
        self._primary = primary
        self._secondary = secondary

    def write(self, data):
        self._primary.write(data)
        self._primary.flush()
        self._secondary.write(data)
        self._secondary.flush()

    def flush(self):
        self._primary.flush()
        self._secondary.flush()

    def fileno(self):
        return self._primary.fileno()


@contextlib.contextmanager
def _cmd_capture():
    log_dir = os.path.join(REPO_ROOT, "repl_logs")
    os.makedirs(log_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    pid = os.getpid()
    log_path = os.path.join(log_dir, f"{pid:010d}_{ts}.log")
    print(f"[capture] log → {log_path}", flush=True)
    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write(f"argv: {sys.argv}\n")
        fh.flush()
        old_out, old_err = sys.stdout, sys.stderr
        sys.stdout = _TeeStream(old_out, fh)
        sys.stderr = _TeeStream(old_err, fh)
        try:
            yield log_path
        finally:
            sys.stdout = old_out
            sys.stderr = old_err


# ---------------------------------------------------------------------------
# Safe imports
# ---------------------------------------------------------------------------

def safe_local_imports(g: dict) -> None:
    try:
        import word_tools
        g["word_tools"] = word_tools
    except Exception:
        import traceback as _tb
        _log(f"FATAL: safe_local_imports failed:\n{_tb.format_exc()}")
        raise


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def _cmd_scan(args: list) -> int:
    path = args[0] if args else DEFAULT_TEMPLATE
    if not os.path.exists(path):
        _log(f"ERROR: file not found: {path}")
        _log(f"       Put the template in: {TEMPLATES_DIR}")
        return 1
    word_tools.scan_document(path)
    return 0


def _next_version_path(base_path: str) -> str:
    """Return the next non-existing versioned path.

    Appends _vNN (zero-padded, two digits) before the extension, incrementing
    until a path that does not yet exist is found.
    Example: 'templates/Form.docx' -> 'templates/Form_v01.docx'
    """
    base, ext = os.path.splitext(base_path)
    v = 1
    while True:
        candidate = f"{base}_v{v:02d}{ext}"
        if not os.path.exists(candidate):
            return candidate
        v += 1


def _cmd_build_template(args: list) -> int:
    """Build a versioned template with Word content controls.

    Copies the read-only source template, opens the copy, then:
      - Inserts plain-text content controls into blank data-entry cells
      - Replaces each \u2610 checkbox character with a checkbox content control
    Saves the result as <basename>_vNN.docx alongside the source template.
    """
    src = args[0] if args else DEFAULT_TEMPLATE
    if not os.path.exists(src):
        _log(f"ERROR: source template not found: {src}")
        return 1

    dst = _next_version_path(src)
    _log(f"Source : {src}")
    _log(f"Output : {dst}")

    word_tools.make_writable_copy(src, dst)

    driver = word_tools.WordDriver(document_path=dst)
    driver.open()
    try:
        # ── Insert text content controls ────────────────────────────────────────────────────────
        text_count = 0
        for tbl, row, col, tag, title, placeholder in _TEXT_CC_FIELDS:
            try:
                rng = driver.table_cell_range(tbl, row, col)
                driver.insert_text_cc(
                    rng, tag, title, placeholder,
                    appearance=word_tools._CC_APPEARANCE_HIDDEN,
                )
                text_count += 1
            except Exception as exc:
                _log(f"  [WARN] text CC ({tbl},{row},{col}) tag={tag!r}: {exc}")
        _log(f"  Inserted {text_count} text content controls")

        # ── Insert paragraph-level text CCs ──────────────────────────────────────────
        para_text_count = 0
        for prefix, tag, title in _PARA_TEXT_CC_FIELDS:
            try:
                rng = driver.clear_para_suffix_for_cc(prefix)
                if rng is not None:
                    driver.insert_text_cc(
                        rng, tag, title,
                        appearance=word_tools._CC_APPEARANCE_HIDDEN,
                    )
                    para_text_count += 1
                else:
                    _log(f"  [WARN] para text CC tag={tag!r}: paragraph not found")
            except Exception as exc:
                _log(f"  [WARN] para text CC tag={tag!r}: {exc}")
        _log(f"  Inserted {para_text_count} paragraph text content controls")

        # ── Insert paragraph-level picture CCs (signatures) ─────────────────────────
        para_pic_count = 0
        for prefix, tag, title in _PARA_PICTURE_CC_FIELDS:
            try:
                rng = driver.clear_para_suffix_for_cc(prefix)
                if rng is not None:
                    driver.insert_picture_cc(rng, tag, title)
                    para_pic_count += 1
                else:
                    _log(f"  [WARN] para picture CC tag={tag!r}: paragraph not found")
            except Exception as exc:
                _log(f"  [WARN] para picture CC tag={tag!r}: {exc}")
        _log(f"  Inserted {para_pic_count} paragraph picture content controls")

        # ── Append device photo CC paragraphs at end of document ─────────────────────
        photo_count = 0
        for tag, title in _DEVICE_PHOTO_TAGS:
            try:
                end_rng = driver._doc.Content
                end_rng.Collapse(0)   # collapse to end
                end_rng.InsertParagraphAfter()
                end_rng.MoveEnd(1, 1)
                end_rng.Collapse(0)
                driver.insert_picture_cc(end_rng, tag, title)
                photo_count += 1
            except Exception as exc:
                _log(f"  [WARN] device photo CC tag={tag!r}: {exc}")
        _log(f"  Inserted {photo_count} device photo content controls")

        # ── Collect all checkbox positions (ascending order) ─────────────────────────
        all_items = []   # (position, tag, title)
        for tbl, row, col, tags_titles in _CELL_CHECKBOXES:
            try:
                cell_rng = driver.table_cell_content_range(tbl, row, col)
                positions = driver.find_char_positions_in_range(cell_rng, "\u2610")
                if len(positions) != len(tags_titles):
                    _log(f"  [WARN] ({tbl},{row},{col}): expected {len(tags_titles)} "
                         f"checkboxes, found {len(positions)} — skipping cell")
                    continue
                for pos, (tag, title) in zip(positions, tags_titles):
                    all_items.append((pos, tag, title))
            except Exception as exc:
                _log(f"  [WARN] checkbox cell ({tbl},{row},{col}): {exc}")

        # Process in descending position order so earlier offsets aren't shifted
        all_items.sort(key=lambda x: x[0], reverse=True)
        cb_count = driver.replace_positions_with_checkbox_ccs(all_items)
        _log(f"  Inserted {cb_count} checkbox content controls")

        driver.save()
        _log(f"Saved: {dst}")
    finally:
        driver.detach()

    return 0


# ---------------------------------------------------------------------------
# Command dispatch
# ---------------------------------------------------------------------------

def _latest_version_path(base_path: str):
    """Return the path of the highest existing vNN version, or None."""
    base, ext = os.path.splitext(base_path)
    found = None
    v = 1
    while True:
        candidate = f"{base}_v{v:02d}{ext}"
        if os.path.exists(candidate):
            found = candidate
            v += 1
        else:
            return found


_SAMPLE_TEXT = {
    "customer_name": "Jane Smith",
    "date_received": "23/05/2026",
    "phone":         "0400 123 456",
    "email":         "jane.smith@example.com",
    "address_notes": "42 Example St, Testville",
    "job_number":    "2026-001",
    "brand_model":   "Dell XPS 15",
    "serial_number": "SN12345678",
    "asset_tag":     "JANE-LAPTOP",
    "received_by":   "Tech 1",
    "time_received": "14:30",
    # Section 9: signature block
    "printed_name":  "Jane Smith",
    "sig_date":      "23/05/2026",
    # Section 11: health-check record
    "tool_profile":  "Nielsoln Health Check v1.0",
    "run_started":   "23/05/2026 14:30",
    "run_completed": "23/05/2026 14:45",
    "report_file":   "2026-001_jane-smith_report.txt",
}

_SAMPLE_CHECKS = {
    "pref_contact_email": True,
    "urgency_normal":     True,
    "device_laptop":      True,
    "os_windows":         True,
    "power_starts":       True,
    "password_no":        True,
    "charger_yes":        True,
}


def _cmd_fill_sample(args: list) -> int:
    """Create a filled sample document from the latest built template version.

    Opens the highest-numbered vNN template, populates the text and checkbox
    fields with representative sample data, and saves the result as
    <basename>_vNN_sample.docx alongside the template.
    """
    src = _latest_version_path(DEFAULT_TEMPLATE)
    if not src:
        _log("ERROR: no built template found; run build_template first")
        return 1
    base, ext = os.path.splitext(src)
    dst = f"{base}_sample{ext}"
    _log(f"Source : {src}")
    _log(f"Sample : {dst}")

    word_tools.make_writable_copy(src, dst)

    driver = word_tools.WordDriver(document_path=dst)
    driver.open()
    try:
        for tag, value in _SAMPLE_TEXT.items():
            ok = driver.set_cc_value(tag, value)
            _log(f"  text  {tag!r}: {'ok' if ok else 'NOT FOUND'}")
        for tag, value in _SAMPLE_CHECKS.items():
            ok = driver.set_cc_value(tag, value)
            _log(f"  check {tag!r}: {'ok' if ok else 'NOT FOUND'}")
        driver.save()
        _log(f"Saved: {dst}")
    finally:
        driver.detach()
    return 0


def _cmd_insert_picture(args: list) -> int:
    """Embed an image file into a named picture content control.

    Usage: insert_picture <tag> <image_path>
    Opens the latest _vNN_sample.docx, finds the CC with the given tag,
    embeds the image, and saves.
    """
    if len(args) < 2:
        _log("Usage: insert_picture <tag> <image_path>")
        return 1
    tag, image_path = args[0], args[1]
    if not os.path.exists(image_path):
        _log(f"ERROR: image not found: {image_path}")
        return 1

    src = _latest_version_path(DEFAULT_TEMPLATE)
    if not src:
        _log("ERROR: no built template found; run build_template first")
        return 1
    base, ext = os.path.splitext(src)
    sample_path = f"{base}_sample{ext}"
    if not os.path.exists(sample_path):
        _log(f"ERROR: sample document not found: {sample_path}")
        _log("       Run fill_sample first to create the sample document.")
        return 1

    driver = word_tools.WordDriver(document_path=sample_path)
    driver.open()
    try:
        ok = driver.set_cc_picture(tag, image_path)
        if ok:
            _log(f"  picture {tag!r}: embedded {os.path.basename(image_path)}")
        else:
            _log(f"  picture {tag!r}: NOT FOUND")
            return 1
        driver.save()
        _log(f"Saved: {sample_path}")
    finally:
        driver.detach()
    return 0


# ---------------------------------------------------------------------------
# Intake field catalogue
# ---------------------------------------------------------------------------
# All tags that onboard_customer understands, grouped by kind.
# text_tags:    populated via set_cc_value(tag, str)
# bool_tags:    populated via set_cc_value(tag, bool)
# picture_tags: populated via set_cc_picture(tag, image_path)

_INTAKE_TEXT_TAGS = [
    # Section 1 — customer details
    "customer_name", "date_received", "phone", "email",
    "address_notes", "job_number",
    # Section 2 — device
    "brand_model", "serial_number", "asset_tag",
    # Section 9 — signature block
    "printed_name", "sig_date",
    # Section 10 — technician intake
    "received_by", "time_received",
    # Section 11 — health-check record
    "tool_profile", "run_started", "run_completed", "report_file",
    # Section 12 — return
    "date_returned",
]

_INTAKE_BOOL_TAGS = [
    # Section 1
    "pref_contact_phone", "pref_contact_text", "pref_contact_email",
    "urgency_normal", "urgency_urgent",
    # Section 2
    "device_laptop", "device_desktop", "device_phone", "device_tablet", "device_other",
    "os_windows", "os_macos", "os_linux", "os_ios", "os_android",
    "power_starts", "power_no_power", "power_intermittent",
    "password_no", "password_yes",
    "charger_no", "charger_yes",
    "acc_bag", "acc_mouse", "acc_keyboard", "acc_usb", "acc_other",
    "damage_no", "damage_yes",
    # Section 10
    "risk_level_low", "risk_level_medium", "risk_level_high",
    "data_risk_low", "data_risk_medium", "data_risk_high",
    "copy_given_yes", "copy_given_no",
]

_INTAKE_PICTURE_TAGS = [
    "sig_customer",        # Section 9 — customer intake signature
    "sig_customer_return", # Section 12 — customer return signature
    "sig_technician",      # Section 12 — technician signature
    "device_front",        # device front photo
    "damage_photo",        # damage/condition photo
]


def _job_slug(job_number: str, customer_name: str) -> str:
    """Return a safe filename slug from job number and customer name."""
    import re
    name_part = re.sub(r"[^a-zA-Z0-9]+", "-", customer_name.strip()).strip("-").lower()
    job_part  = re.sub(r"[^a-zA-Z0-9]+", "-", job_number.strip()).strip("-")
    return f"{job_part}_{name_part}" if name_part else job_part


def _cmd_next_job(args: list) -> int:
    """Claim and print the next job number.

    If GITHUB_TOKEN and GITHUB_REPO are configured in locals.txt, the claim is
    race-protected via the GitHub Contents API (optimistic locking).  Prints
    the claimed job ID so it can be copied into the intake JSON.
    """
    try:
        job_id = _next_job_number()
        _log(f"Job number: {job_id}")
        return 0
    except Exception as exc:
        _log(f"ERROR: could not claim job number: {exc}")
        return 1


_CUSTOMERS_TSV_COLUMNS = ["customer_id", "customer_name", "phone", "email", "address_notes"]
_JOBS_TSV_COLUMNS = [
    "job_number", "customer_id", "date_received", "brand_model",
    "serial_number", "asset_tag", "received_by", "time_received",
]


def _lookup_customer(data: dict):
    """Look up a customer in customers.tsv by name, phone, and email.

    Returns (customer_id, conflicts) where:
      customer_id: str if all three fields match an existing row, else None
      conflicts:   list of rows where 1 or 2 of the 3 fields match
    """
    import csv
    name  = str(data.get("customer_name", "")).strip().lower()
    phone = str(data.get("phone", "")).strip()
    email = str(data.get("email", "")).strip().lower()

    if not os.path.exists(CUSTOMERS_TSV) or os.path.getsize(CUSTOMERS_TSV) == 0:
        return None, []

    conflicts = []
    with open(CUSTOMERS_TSV, encoding="utf-8", newline="") as fh:
        for row in csv.DictReader(fh, delimiter="\t"):
            row_name  = row.get("customer_name", "").strip().lower()
            row_phone = row.get("phone", "").strip()
            row_email = row.get("email", "").strip().lower()
            score = (name == row_name) + (phone == row_phone) + (email == row_email)
            if score == 3:
                return row["customer_id"], []
            if score >= 1:
                conflicts.append(row)
    return None, conflicts


def _append_customers_tsv(customer_id: str, data: dict) -> None:
    """Append one row to customers.tsv for a newly created customer."""
    import csv
    row = {k: data.get(k, "") for k in _CUSTOMERS_TSV_COLUMNS}
    row["customer_id"] = customer_id
    write_header = not os.path.exists(CUSTOMERS_TSV) or os.path.getsize(CUSTOMERS_TSV) == 0
    with open(CUSTOMERS_TSV, "a", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=_CUSTOMERS_TSV_COLUMNS, delimiter="\t")
        if write_header:
            writer.writeheader()
        writer.writerow(row)
    _log(f"  customers.tsv: {customer_id} added ({data.get('customer_name', '?')})")


def _append_jobs_tsv(customer_id: str, data: dict) -> None:
    """Append one row to jobs.tsv for the just-created job."""
    import csv
    row = {k: data.get(k, "") for k in _JOBS_TSV_COLUMNS}
    row["customer_id"] = customer_id
    write_header = not os.path.exists(JOBS_TSV) or os.path.getsize(JOBS_TSV) == 0
    with open(JOBS_TSV, "a", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=_JOBS_TSV_COLUMNS, delimiter="\t")
        if write_header:
            writer.writeheader()
        writer.writerow(row)
    _log(f"  jobs.tsv: row appended for job {data.get('job_number', '?')}")


def _cmd_onboard_customer(args: list) -> int:
    """Fill a new intake form from a JSON data file.

    Usage: onboard_customer <json_path>

    The JSON file must be a flat object whose keys are CC tags.  Text and
    bool values are set directly; picture tags must map to absolute image
    file paths.  Unrecognised keys are logged and skipped.

    Output: jobs/<slug>.docx  where slug = <job_number>_<customer_name>.
    """
    import json

    if not args:
        _log("Usage: onboard_customer <json_path>")
        return 1
    json_path = args[0]
    if not os.path.exists(json_path):
        _log(f"ERROR: JSON file not found: {json_path}")
        return 1

    with open(json_path, encoding="utf-8") as fh:
        data = json.load(fh)

    src = _latest_version_path(DEFAULT_TEMPLATE)
    if not src:
        _log("ERROR: no built template found; run build_template first")
        return 1

    job_number    = str(data.get("job_number", "job")).strip()
    customer_name = str(data.get("customer_name", "customer")).strip()
    slug = _job_slug(job_number, customer_name)

    # ── Customer dedup check (before creating any files) ─────────────────────
    customer_id, conflicts = _lookup_customer(data)
    if conflicts:
        _log("  WARNING: partial customer match — manual review required:")
        for c in conflicts:
            _log(f"    {c}")
        _log("  Onboarding aborted. Resolve customer identity before proceeding.")
        return 1

    os.makedirs(JOBS_DIR, exist_ok=True)
    dst = os.path.join(JOBS_DIR, f"{slug}.docx")
    _log(f"Template : {src}")
    _log(f"Output   : {dst}")

    word_tools.make_writable_copy(src, dst)

    driver = word_tools.WordDriver(document_path=dst)
    driver.open()
    try:
        known_text = set(_INTAKE_TEXT_TAGS)
        known_bool = set(_INTAKE_BOOL_TAGS)
        known_pic  = set(_INTAKE_PICTURE_TAGS)

        for tag, value in data.items():
            if tag in known_text:
                ok = driver.set_cc_value(tag, str(value))
                _log(f"  text    {tag!r}: {'ok' if ok else 'NOT FOUND'}")
            elif tag in known_bool:
                ok = driver.set_cc_value(tag, bool(value))
                _log(f"  check   {tag!r}: {'ok' if ok else 'NOT FOUND'}")
            elif tag in known_pic:
                img = str(value)
                if os.path.exists(img):
                    ok = driver.set_cc_picture(tag, img)
                    _log(f"  picture {tag!r}: {'embedded' if ok else 'CC NOT FOUND'}")
                else:
                    _log(f"  picture {tag!r}: SKIP (file not found: {img})")
            else:
                _log(f"  SKIP    {tag!r}: unrecognised tag")

        driver.save()
        _log(f"Saved: {dst}")
    finally:
        driver.detach()

    # ── Write TSVs ────────────────────────────────────────────────────────────
    if customer_id:
        _log(f"  customers.tsv: existing customer {customer_id}")
    else:
        customer_id = _next_customer_id()
        _append_customers_tsv(customer_id, data)
    _append_jobs_tsv(customer_id, data)
    return 0


def _cmd_capture_sig(args: list) -> int:
    """Capture a handwritten signature and embed it into a job document.

    Usage: capture_sig <tag> <job_doc_path>

    Opens a maximised drawing canvas.  The customer signs with the Wacom pen or
    mouse.  "Clear" resets the canvas.  "Cancel" closes without saving.
    "Accept" saves the strokes as a PNG (temp file) and embeds it into the
    named picture CC in the job document.

    Requires Pillow (auto-installed on first run).
    """
    import ctypes
    import ctypes.wintypes
    import tempfile
    import tkinter as tk

    if len(args) < 2:
        _log("Usage: capture_sig <tag> <job_doc_path>")
        return 1
    tag, doc_path = args[0], args[1]
    if not os.path.exists(doc_path):
        _log(f"ERROR: document not found: {doc_path}")
        return 1

    PIL_Image     = _install_and_import("PIL.Image",     "Pillow")
    PIL_ImageDraw = _install_and_import("PIL.ImageDraw", "Pillow")

    result = {"path": None}

    # ── ClipCursor helpers — confine tablet/mouse to the canvas screen rect ─
    def _lock_cursor_to_canvas():
        """Restrict all cursor movement to the full Tkinter window."""
        root.update_idletasks()
        x = root.winfo_rootx()
        y = root.winfo_rooty()
        rect = ctypes.wintypes.RECT(x, y, x + root.winfo_width(),
                                    y + root.winfo_height())
        ctypes.windll.user32.ClipCursor(ctypes.byref(rect))

    def _unlock_cursor():
        """Remove cursor confinement."""
        ctypes.windll.user32.ClipCursor(None)

    # ── Bootstrap root to measure maximised dimensions ──────────────────────
    root = tk.Tk()
    root.title(f"Sign here  —  {tag}")
    root.state("zoomed")   # maximise (keeps title bar + taskbar)
    root.bind_class("Button", "<Return>", lambda e: e.widget.invoke())
    root.update()          # process geometry so winfo_width/height are valid

    BTN_H = 52             # approximate height of the button row
    W = root.winfo_width()
    H = root.winfo_height() - BTN_H

    pil_img  = PIL_Image.new("RGB", (W, H), "white")
    pil_draw = PIL_ImageDraw.Draw(pil_img)

    canvas = tk.Canvas(root, width=W, height=H, bg="white", cursor="crosshair",
                       highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)

    # Dashed baseline — sits 60 px above the bottom of the canvas
    canvas.create_line(40, H - 60, W - 40, H - 60, fill="#cccccc", dash=(6, 4))

    # Lock cursor once the canvas is fully laid out
    root.after(100, _lock_cursor_to_canvas)

    prev = {}

    def on_press(event):
        prev["x"], prev["y"] = event.x, event.y

    def on_drag(event):
        x0, y0 = prev.get("x", event.x), prev.get("y", event.y)
        x1, y1 = event.x, event.y
        canvas.create_line(x0, y0, x1, y1, width=3, fill="black",
                           capstyle=tk.ROUND, joinstyle=tk.ROUND)
        pil_draw.line([(x0, y0), (x1, y1)], fill="black", width=3)
        prev["x"], prev["y"] = x1, y1

    def on_clear():
        canvas.delete("all")
        canvas.create_line(40, H - 60, W - 40, H - 60, fill="#cccccc", dash=(6, 4))
        pil_draw.rectangle([0, 0, W, H], fill="white")

    def on_cancel():
        _unlock_cursor()
        root.destroy()

    def on_accept():
        _unlock_cursor()
        # Crop to the bounding box of non-white (ink) pixels, with a small margin.
        gray = pil_img.convert("L")
        # Invert relative to threshold: white bg → 0, ink strokes → 255
        mask = gray.point(lambda p: 0 if p > 240 else 255)
        bbox = mask.getbbox()
        if bbox:
            pad = 12
            iw, ih = pil_img.size
            bbox = (max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                    min(iw, bbox[2] + pad), min(ih, bbox[3] + pad))
            img_to_save = pil_img.crop(bbox)
        else:
            img_to_save = pil_img   # blank canvas — save as-is
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.close()
        img_to_save.save(tmp.name)
        result["path"] = tmp.name
        root.destroy()

    # Safety net: always unlock if the window is destroyed by any other means
    root.bind("<Destroy>", lambda _e: _unlock_cursor())

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>",     on_drag)

    btn_frame = tk.Frame(root, pady=4)
    btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Button(btn_frame, text="Cancel", command=on_cancel, width=12,
              bg="#e53935", fg="white", font=("Segoe UI", 11)).pack(
        side=tk.LEFT, padx=12)
    tk.Button(btn_frame, text="Clear",  command=on_clear,  width=12,
              font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=4)
    tk.Button(btn_frame, text="Accept", command=on_accept, width=12,
              bg="#4caf50", fg="white", font=("Segoe UI", 11, "bold")).pack(
        side=tk.RIGHT, padx=12)

    root.mainloop()

    if not result["path"]:
        _log("  Signature capture cancelled (window closed without Accept).")
        return 0

    _log(f"  Signature PNG: {result['path']}")

    driver = word_tools.WordDriver(document_path=doc_path)
    driver.open()
    try:
        ok = driver.insert_picture_at_tag(tag, result["path"])
        if ok:
            _log(f"  picture {tag!r}: embedded")
        else:
            _log(f"  picture {tag!r}: CC/bookmark NOT FOUND — PNG at {result['path']}")
            return 1
        driver.save()
        _log(f"Saved: {doc_path}")
    finally:
        driver.detach()

    return 0


COMMANDS = {
    "scan":              _cmd_scan,
    "build_template":    _cmd_build_template,
    "fill_sample":       _cmd_fill_sample,
    "insert_picture":    _cmd_insert_picture,
    "next_job":          _cmd_next_job,
    "onboard_customer":  _cmd_onboard_customer,
    "capture_sig":       _cmd_capture_sig,
}


def main() -> int:
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    args = sys.argv[2:]

    if not cmd:
        _log("Commands: " + ", ".join(sorted(COMMANDS)))
        _log("Usage:    sys.argv[1:] = ['<command>'] ; import runpy ; temp = runpy._run_module_as_main('repl_main')")
        return 0

    handler = COMMANDS.get(cmd)
    if handler is None:
        _log(f"ERROR: unknown command '{cmd}'. Known: {', '.join(sorted(COMMANDS))}")
        return 1

    return handler(args)


if __name__ == "__main__":
    safe_local_imports(globals())
    with _cmd_capture():
        try:
            rc = main()
            if rc:
                _log(f"main() returned {rc}")
        except Exception:
            import traceback
            _log(f"FATAL unhandled exception:\n{traceback.format_exc()}")
            raise
