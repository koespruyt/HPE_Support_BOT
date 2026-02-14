import argparse
import csv
import json
import re
import subprocess
import sys

# Avoid UnicodeEncodeError on Windows consoles
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from pathlib import Path
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


def _force_utf8_stdio() -> None:
    """Prevent UnicodeEncodeError on Windows consoles/scheduled tasks.

    - When stdout/stderr are cp1252, printing emoji/Unicode can crash.
    - Reconfigure to UTF-8 + replace errors so the script keeps running.
    """
    for s in (sys.stdout, sys.stderr):
        try:
            if hasattr(s, "reconfigure"):
                s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


_force_utf8_stdio()

# --- Field label mapping (UI text -> normalized key) ---
# If HPE changes UI labels, update these lists or hpe_selectors.json (preferred).
FIELD_LABELS: Dict[str, List[str]] = {
    # case detail panel labels (most common)
    'status': ['Status', 'Case Status'],
    'severity': ['Severity'],
    'product': ['Product'],
    'serial': ['Serial Number', 'Serial number', 'Serial'],
    'product_no': ['Product Number', 'Product number'],
    'group': ['Group'],
    'nickname': ['Nickname'],
    # sometimes shown in header area
    'case_no': ['Case', 'Case:', 'Case Number', 'Case number'],
}


def utc_now_iso():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))

def save_text(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text or "", encoding="utf-8", errors="ignore")

def save_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")

def redact_sensitive(text: str) -> str:
    """Redact obvious secrets from Communications export (passwords, tokens).

    Keeps logins/usernames and URLs, but replaces any password/token values.
    """
    if not text:
        return ""
    t = text

    # Line-based (handles indentation)
    t = re.sub(r"(?im)^\s*(Password\s*:\s*).+$", r"\1[REDACTED]", t)
    t = re.sub(r"(?im)^\s*(Password\s*Token\s*:\s*).+$", r"\1[REDACTED]", t)
    t = re.sub(r"(?im)^\s*(Wrap\s*token\s*:\s*).+$", r"\1[REDACTED]", t)
    t = re.sub(r"(?im)^\s*(Token\s*:\s*).+$", r"\1[REDACTED]", t)

    # Inline fallbacks (covers 'Password: xyz' in the middle of a line)
    t = re.sub(r"(?i)(Password\s*:\s*)([^\s\r\n]+)", r"\1[REDACTED]", t)
    t = re.sub(r"(?i)(Password\s*Token\s*:\s*)([^\s\r\n]+)", r"\1[REDACTED]", t)
    t = re.sub(r"(?i)(Wrap\s*token\s*:\s*)([^\s\r\n]+)", r"\1[REDACTED]", t)

    return t

def first_visible_locator(page, selector_list):
    for sel in selector_list:
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return None

def any_text_present(page, texts):
    for s in texts:
        try:
            if page.locator(f"text={s}").count() > 0:
                return True
        except Exception:
            continue
    return False

def ensure_ready(page, cfg, timeout_ms=45000):
    start = datetime.now()
    while (datetime.now() - start).total_seconds() * 1000 < timeout_ms:
        if any_text_present(page, cfg.get("session_expired_text_any", [])):
            raise RuntimeError("SESSION_EXPIRED")
        if any_text_present(page, cfg.get("ready_text_any", [])):
            return True
        if page.locator("text=/\\bCase\\s+\\d{7,12}\\b/").count() > 0:
            return True
        page.wait_for_timeout(500)
    raise RuntimeError("CASES_PAGE_NOT_READY_TIMEOUT")

def dismiss_contract_banner(page, cfg):
    if not any_text_present(page, cfg.get("contract_banner_text_any", [])):
        return False
    btn = first_visible_locator(page, cfg.get("contract_banner_dismiss_any", []))
    if btn:
        try:
            btn.click(timeout=5000)
            page.wait_for_timeout(500)
            return True
        except Exception:
            return False
    return False

MONTHS = r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
TS_RX = re.compile(rf"\b{MONTHS}\s+\d{{1,2}},\s+\d{{4}},\s+\d{{1,2}}:\d{{2}}\s+(AM|PM)\b")
URL_RX = re.compile(r"https?://[^\s)>\"]+", re.I)
HOST_LINE_RX = re.compile(r"(?i)^\s*(System Name/Host Name|System Name|Host Name)\s*:\s*(.*?)\s*$")
HOST_TOKEN_RX = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]{2,}")
STOP_HOST_VALUES = {"problem", "additional", "serial", "case", "event", "none", "null", "n/a"}
DEAR_RX = re.compile(r"(?im)^\s*Dear\s+(.+?),\s*$")
ADDR_MARKERS = ["Equipment Address", "Site Address", "Customer Address", "Address"]
ADDR_KV_RX = re.compile(r"(?im)^\s*(Street|City|State|Province|Postal Code|Postcode|Zip|ZIP|Country)\s*:\s*(.+?)\s*$")

def parse_line_pairs(text: str) -> Dict[str, str]:
    """Parse label/value pairs from blocks where a label is followed by a value.

    Also supports 'Label: value' on a single line.
    """
    if not text:
        return {}
    lines = [ln.strip() for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    out: Dict[str, str] = {}

    # For quick lookup, map lower(label) -> field key
    label_map: Dict[str, str] = {}
    for key, labels in FIELD_LABELS.items():
        for lbl in labels:
            label_map[lbl.lower()] = key

    for i, ln in enumerate(lines):
        if not ln:
            continue

        # Case: "Label: value" on same line
        if ":" in ln:
            left, right = ln.split(":", 1)
            k = label_map.get(left.strip().lower())
            if k and right.strip():
                out.setdefault(k, right.strip())

        key = label_map.get(ln.lower())
        if not key:
            continue

        # Next non-empty line is value
        j = i + 1
        while j < len(lines) and not lines[j]:
            j += 1
        if j >= len(lines):
            continue
        val = lines[j]
        if not val or val in ("-", "—"):
            continue
        # Avoid capturing another label as value
        if val.lower() in label_map:
            continue
        out.setdefault(key, val)

    return out

def split_messages(text: str):
    if not text:
        return []
    norm = text.replace("\r\n", "\n").replace("\r", "\n")
    hits = [(m.start(), m.group(0)) for m in TS_RX.finditer(norm)]
    if not hits:
        return []
    blocks = []
    for idx, (pos, ts) in enumerate(hits):
        end = hits[idx + 1][0] if idx + 1 < len(hits) else len(norm)
        block = norm[pos:end].strip()
        blocks.append((ts, block))
    return blocks

def extract_subject(block: str) -> str:
    if not block:
        return ""
    lines = [l.strip() for l in block.splitlines()]
    for i, l in enumerate(lines):
        if l.lower() == "subject":
            for j in range(i + 1, min(i + 12, len(lines))):
                if lines[j]:
                    return lines[j]
    return ""

def parse_ts(ts: str):
    for fmt in ("%b %d, %Y, %I:%M %p", "%b %d, %Y %I:%M %p"):
        try:
            return datetime.strptime(ts.strip(), fmt)
        except Exception:
            pass
    return None

def pick_last_hpe_message(blocks):
    best = None
    best_dt = None
    for ts, block in blocks:
        dt = parse_ts(ts)
        low = block.lower()
        is_hpe = ("hpe support engineer" in low) or ("hewlett packard enterprise" in low) or ("hpe services" in low)
        has_subject = ("subject" in low)
        if not (is_hpe and has_subject):
            continue
        if dt and (best_dt is None or dt > best_dt):
            best_dt = dt
            best = (ts, block)
    if best is None and blocks:
        scored = []
        for ts, block in blocks:
            dt = parse_ts(ts) or datetime.min
            scored.append((dt, ts, block))
        scored.sort(reverse=True)
        _, ts, block = scored[0]
        best = (ts, block)
    return best

def extract_key_links(text: str, limit=10):
    """Extract and rank URLs from Communications text."""
    if not text:
        return []
    rx = re.compile(r"https?://[^\s)\]}>\"']+", re.I)
    seen = []
    for m in rx.finditer(text):
        url = m.group(0).strip().rstrip(".,;\"'")
        if url not in seen:
            seen.append(url)

    def rank(u: str) -> int:
        lu = u.lower()
        if "hprc" in lu:
            return 0
        if "scts.it.hpe.com" in lu:
            return 1
        if "ahscatalogsearch" in lu:
            return 2
        if "support.hpe.com" in lu:
            return 3
        return 4

    seen.sort(key=rank)
    return seen[:limit]



def extract_event_ids(text: str):
    if not text:
        return []
    rx = re.compile(r"\bEvent Id:\s*([0-9a-fA-F-]{36})\b")
    seen = []
    for m in rx.finditer(text):
        v = m.group(1)
        if v not in seen:
            seen.append(v)
    return seen


def extract_problem_descriptions(text: str, limit: int = 5):
    if not text:
        return []
    rx = re.compile(r"\bProblem Description:\s*([^\r\n]{3,300})", re.I)
    out = []
    for m in rx.finditer(text):
        v = m.group(1).strip()
        if v and v not in out:
            out.append(v)
        if len(out) >= limit:
            break
    return out


def extract_ahs_links(text: str, limit: int = 5):
    if not text:
        return []
    rx = re.compile(r"https?://ahscatalogsearch\.it\.hpe\.com/\?log=[^\s)\]}>\"']+", re.I)
    out = []
    for m in rx.finditer(text):
        u = m.group(0).strip().rstrip(".,;\"'")
        if u not in out:
            out.append(u)
        if len(out) >= limit:
            break
    return out


def extract_dropbox_info(text: str):
    """Return (dropbox_hosts, logins) from HPRC instructions."""
    if not text:
        return ([], [])
    hosts = []
    logins = []
    for u in extract_key_links(text, limit=50):
        lu = u.lower()
        if "hprc" in lu or "hprc-h" in lu:
            # normalize host
            m = re.match(r"https?://([^/]+)/?", u, re.I)
            if m:
                h = m.group(1)
                if h not in hosts:
                    hosts.append(h)
    rx_login = re.compile(r"\bLogin:\s*([A-Za-z0-9._-]{3,32})\b", re.I)
    for m in rx_login.finditer(text):
        v = m.group(1).strip()
        if v not in logins:
            logins.append(v)
    return (hosts, logins)

def find_host_name(text: str) -> str:
    """Try to extract a hostname/system name from the Communications text.

    The portal sometimes renders the 'System Name/Host Name' line without clean newlines,
    so we do a full-text search rather than relying purely on line starts.
    """
    if not text:
        return ""

    m = re.search(r"(?is)(System\s*Name/Host\s*Name|Host\s*Name|System\s*Name)\s*:\s*([^\r\n]{0,200})", text)
    if m:
        val = (m.group(2) or "").strip()
        for cut in ["You will", "You can", "You may", "You should", "You will receive", "You can view"]:
            if cut in val:
                val = val.split(cut, 1)[0].strip()
        tm = HOST_TOKEN_RX.search(val)
        if tm:
            host = tm.group(0).strip()
            # Sometimes the UI glues words right after the hostname (e.g. "src123YouWill...").
            # If the token ends with letters after the last digit, trim it back to the last digit.
            m2 = re.match(r"^([A-Za-z0-9._-]*\d)[A-Za-z]{2,}$", host)
            if m2:
                host = m2.group(1)
            if host.lower() not in STOP_HOST_VALUES:
                return host

    # Fallback: line-based
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        mm = HOST_LINE_RX.match(line.strip())
        if not mm:
            continue
        val = (mm.group(2) or "").strip()
        if not val:
            continue
        for cut in ["You will", "You can", "You may", "You should"]:
            if cut in val:
                val = val.split(cut, 1)[0].strip()
        tm = HOST_TOKEN_RX.search(val)
        if not tm:
            continue
        host = tm.group(0).strip()
        if host.lower() in STOP_HOST_VALUES:
            continue
        return host

    return ""

def find_salutation_name(text: str) -> str:
    m = DEAR_RX.search(text or "")
    return (m.group(1).strip() if m else "")

def extract_address_block(text: str) -> dict:
    if not text:
        return {}
    lines = [l.strip() for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]

    def is_marker(line: str) -> bool:
        low = line.lower().rstrip(":")
        return any(low == m.lower().rstrip(":") for m in ADDR_MARKERS) or any(low.startswith(m.lower()) for m in ADDR_MARKERS)

    for i, line in enumerate(lines):
        if is_marker(line):
            d = {}
            for j in range(i + 1, min(i + 30, len(lines))):
                if not lines[j]:
                    if d:
                        break
                    continue
                stop_low = lines[j].lower()
                if stop_low.startswith(("thank you", "sincerely", "ref:")):
                    break
                m = ADDR_KV_RX.match(lines[j])
                if m:
                    key = m.group(1).strip().lower()
                    val = m.group(2).strip()
                    if key in ("zip", "zipcode"):
                        key = "postal_code"
                    if key in ("postal code", "postcode"):
                        key = "postal_code"
                    if key == "province":
                        key = "state"
                    d[key] = val
            if d:
                return d
    return {}

def infer_requested_actions(status: str, comms: str):
    st = (status or "").lower()
    c = (comms or "").lower()

    actions = []
    category = "UNKNOWN"
    summary = ""

    if "approve" in st and "closure" in st:
        category = "CLOSE_APPROVAL"
        summary = "HPE wacht op goedkeuring om case te sluiten."
        actions.append("Bevestigen dat de case opgelost is en mag afgesloten worden (Approve case closure).")
        return category, summary, actions

    if "complete action plan" in st:
        category = "ACTION_PLAN"
        summary = "HPE wacht op completion van het action plan."
        actions.append("Action plan afronden en bevestigen in HPE portal (Complete action plan).")

    if ("log file request" in c) or ("require some log files" in c) or ("provide these logs" in c) or ("active health system" in c) or ("ahs log" in c):
        category = "LOG_REQUEST"
        summary = "HPE vraagt logbestanden (AHS) + eventueel iLO/disk info om diagnose verder te zetten."
        actions.append("AHS / Active Health System log genereren en uploaden naar HPE dropbox (HPRC).")
        if "reply all" in c:
            actions.append("Na upload: Reply All op HPE mail/thread om te bevestigen dat logs klaarstaan.")
        if ("ilo storage" in c) or ("photo of the physical drive" in c) or ("led" in c):
            actions.append("Screenshot iLO Storage tab of foto fysieke disk/LEDs toevoegen.")

    if not actions and "awaiting customer action" in st:
        category = "AWAITING"
        summary = "Case staat op Awaiting Customer Action."
        actions.append("Customer action vereist: check laatste HPE communicatie voor exacte vraag.")

    if not actions:
        summary = "Onvoldoende signalen om vraag automatisch te bepalen."
        actions.append("Manuele check Communications nodig.")

    return category, summary, actions

def collect_case_numbers(page, cfg, max_cases: int):
    container = first_visible_locator(page, cfg.get("case_list_container_any", []))
    scope = container if container else page
    rx = re.compile(cfg.get("case_text_regex", r"\bCase\s+(\d{7,12})\b"))
    found, seen = [], set()

    for _ in range(15):
        try:
            loc = scope.locator("text=/\\bCase\\s+\\d{7,12}\\b/")
            texts = loc.all_text_contents()
        except Exception:
            texts = []

        for t in texts:
            m = rx.search(t)
            if m:
                cn = m.group(1)
                if cn not in seen:
                    seen.add(cn)
                    found.append(cn)
                    if max_cases > 0 and len(found) >= max_cases:
                        return found

        try:
            if container:
                container.evaluate("el => { el.scrollTop = el.scrollHeight; }")
            else:
                page.mouse.wheel(0, 1200)
        except Exception:
            pass

        page.wait_for_timeout(800)

    return found

def find_case_search_input(page, cfg):
    for sel in cfg.get("case_search_input_any", []):
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return None

def open_case_by_number(page, cfg, case_no: str):
    search = find_case_search_input(page, cfg)
    container = first_visible_locator(page, cfg.get("case_list_container_any", []))
    scope = container if container else page

    if search:
        try:
            search.click(timeout=5000)
            search.fill("")
            search.fill(case_no)
            page.keyboard.press("Enter")
            page.wait_for_timeout(600)
        except Exception:
            pass

    target = scope.locator(f"text=/\\bCase\\s+{case_no}\\b/").first
    try:
        target.wait_for(timeout=15000)
        target.click(timeout=15000)
    except Exception as e:
        raise RuntimeError(f"Could not click case in list: {case_no} ({e})")

    try:
        page.locator(f"text=/\\bCase\\s*:?\\s*{case_no}\\b/").first.wait_for(timeout=25000)
    except PWTimeoutError:
        pass

def click_tab(page, cfg, tab_key: str):
    sels = cfg.get(tab_key, [])
    for sel in sels:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0:
                loc.click(timeout=8000)
                page.wait_for_timeout(400)
                return True
        except Exception:
            continue
    return False

def extract_details_text(page):
    panels = page.get_by_role("tabpanel")
    try:
        count = panels.count()
    except Exception:
        count = 0

    for i in range(min(count, 8)):
        p = panels.nth(i)
        try:
            if p.is_visible():
                txt = p.inner_text(timeout=5000)
                if "Serial Number" in txt and "Status" in txt:
                    return txt
        except Exception:
            continue

    try:
        return page.locator("main").inner_text(timeout=5000)
    except Exception:
        return page.content()

def click_expand_all_comms(page, cfg):
    for sel in cfg.get("expand_all_any", []):
        try:
            loc = page.locator(sel).first
            if loc.count() == 0:
                continue
            loc.click(timeout=8000)
            page.wait_for_timeout(500)
            return True
        except Exception:
            continue
    return False



def ensure_expand_all_comms(page, cfg) -> bool:
    """Try to ensure 'Expand all communications' is enabled (checkbox/toggle)."""
    try:
        cb = page.get_by_role("checkbox", name=re.compile(r"Expand all communications", re.I))
        if cb.count():
            try:
                if not cb.first.is_checked():
                    cb.first.set_checked(True, timeout=8000)
                    page.wait_for_timeout(500)
            except Exception:
                cb.first.click(timeout=8000)
                page.wait_for_timeout(500)
            return True
    except Exception:
        pass
    return click_expand_all_comms(page, cfg)


def click_all_read_more(page, cfg, max_rounds: int = 10):
    """Click all visible 'Read more' links/buttons inside Communications."""
    selectors = cfg.get("read_more_any") or [
        "a:has-text('Read more')",
        "button:has-text('Read more')",
        "text=/Read more/i",
    ]
    for _ in range(max_rounds):
        clicked = False
        for sel in selectors:
            try:
                loc = page.locator(sel)
                n = loc.count()
                if n == 0:
                    continue
                for i in range(min(n, 30)):
                    try:
                        item = loc.nth(i)
                        if item.is_visible():
                            item.click(timeout=3000, force=True)
                            page.wait_for_timeout(250)
                            clicked = True
                    except Exception:
                        continue
            except Exception:
                continue
        if not clicked:
            break

def extract_comms_text(page, cfg):
    panels = page.get_by_role("tabpanel")
    try:
        count = panels.count()
    except Exception:
        count = 0

    hints = cfg.get("comms_panel_hint_any", [])

    for i in range(min(count, 10)):
        p = panels.nth(i)
        try:
            if not p.is_visible():
                continue
            txt = p.inner_text(timeout=8000)
            if any(h in txt for h in hints):
                return txt
        except Exception:
            continue

    try:
        return page.locator("main").inner_text(timeout=8000)
    except Exception:
        return ""

def write_alarm(outdir: Path, alarm_file: str, msg: str, alarm_cmd: str | None):
    p = outdir / alarm_file
    save_text(p, f"[{utc_now_iso()}] {msg}\\n")
    if alarm_cmd:
        try:
            subprocess.run(alarm_cmd, shell=True, check=False)
        except Exception:
            pass

def main():
    ap = argparse.ArgumentParser(description="HPE Support Center - Cases overview + communications extraction (standalone)")
    ap.add_argument("--state", default="hpe_state.json", help="Playwright storage state JSON (logged-in session)")
    ap.add_argument("--selectors", default="hpe_selectors.json", help="Selectors/config JSON (easy GUI tweaks)")
    ap.add_argument("--outdir", default="out_hpe", help="Output directory")
    ap.add_argument("--max", type=int, default=0, help="Max cases to process (0=all found)")
    ap.add_argument("--headless", action="store_true", help="Run Chromium headless")
    ap.add_argument("--format", default="both", choices=["csv", "json", "both"], help="Export format for overview")
    ap.add_argument("--alarm-file", default="ALERT_SESSION_EXPIRED.txt", help="Alarm file to write on session timeout")
    ap.add_argument("--alarm-cmd", default=None, help="Optional command to execute when session expired (e.g. PowerShell mail)")
    args = ap.parse_args()

    state_path = Path(args.state)
    cfg_path = Path(args.selectors)
    outdir = Path(args.outdir)
    (outdir / "cases").mkdir(parents=True, exist_ok=True)
    (outdir / "debug").mkdir(parents=True, exist_ok=True)

    if not state_path.exists():
        print(f"ERROR: Missing state file: {state_path}")
        print("Tip: run python .\\01_login_save_state.py to generate hpe_state.json")
        return 3

    if not cfg_path.exists():
        print(f"ERROR: Missing selectors file: {cfg_path}")
        return 3

    cfg = load_json(cfg_path)
    cases_url = cfg.get("cases_url", "https://support.hpe.com/connect/s/?tab=cases")
    # Avoid emoji/unicode in console output (Windows codepages can mangle it)
    print(f"Open cases: {cases_url}")

    results, errors = [], []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(storage_state=str(state_path))
        page = context.new_page()

        try:
            page.goto(cases_url, wait_until="domcontentloaded", timeout=60000)
            dismiss_contract_banner(page, cfg)

            try:
                ensure_ready(page, cfg, timeout_ms=45000)
            except RuntimeError as e:
                if str(e) == "SESSION_EXPIRED":
                    write_alarm(outdir, args.alarm_file, "Session expired / login required on HPE portal.", args.alarm_cmd)
                    print("SESSION_EXPIRED (alarm written)")
                    return 2
                raise

            case_nos = collect_case_numbers(page, cfg, args.max)
            case_nos = case_nos[:args.max] if args.max > 0 else case_nos
            case_nos = [c for c in case_nos if c]
            print(f"Found cases: {len(case_nos)} -> " + ", ".join(case_nos[:20]) + (" ..." if len(case_nos) > 20 else ""))

            for idx, case_no in enumerate(case_nos, start=1):
                print(f"\n=== [{idx}/{len(case_nos)}] Case {case_no} ===")
                try:
                    open_case_by_number(page, cfg, case_no)

                    click_tab(page, cfg, "tab_details_any")
                    page.wait_for_timeout(500)
                    details_text = redact_sensitive(extract_details_text(page))
                    fields = parse_line_pairs(details_text)
                    serial = fields.get("serial", "") or ""

                    click_tab(page, cfg, "tab_communications_any")
                    page.wait_for_timeout(700)
                    click_expand_all_comms(page, cfg)
                    page.wait_for_timeout(900)
                    comms_text = redact_sensitive(extract_comms_text(page, cfg))
                    # Fallback: if Overview/Details fields are missing, parse from Communications header
                    fields_from_comms = parse_line_pairs(comms_text)
                    for k, v in fields_from_comms.items():
                        if v and not fields.get(k):
                            fields[k] = v

                    comms_file = outdir / "cases" / f"{case_no}_communications_redacted.txt"
                    save_text(comms_file, comms_text)

                    blocks = split_messages(comms_text)
                    last = pick_last_hpe_message(blocks)
                    last_ts, last_block = (last if last else ("", ""))

                    hpe_subject = extract_subject(last_block)
                    host = find_host_name(comms_text)
                    contact_name = find_salutation_name(comms_text)
                    addr = extract_address_block(comms_text)

                    # Extra structured info from communications (handy for dashboards)
                    event_ids = extract_event_ids(comms_text)
                    problem_descs = extract_problem_descriptions(comms_text)
                    ahs_links = extract_ahs_links(comms_text)
                    drop_hosts, drop_logins = extract_dropbox_info(comms_text)

                    category, summary, actions = infer_requested_actions(fields.get("status",""), comms_text)
                    key_links = extract_key_links(last_block or comms_text, limit=8)

                    action_plan = fields.get("action_plan", "")
                    if not action_plan and fields.get("status"):
                        sl = fields["status"].lower()
                        if "complete action plan" in sl:
                            action_plan = "Complete action plan"
                        elif "approve case closure" in sl or "approve closure" in sl:
                            action_plan = "Approve case closure"
                    row = {
                        "case_no": case_no,
                        "serial": serial,
                        "host_name": host,
                        "contact_name": contact_name,
                        "addr_street": addr.get("street",""),
                        "addr_city": addr.get("city",""),
                        "addr_state": addr.get("state",""),
                        "addr_postal_code": addr.get("postal_code",""),
                        "addr_country": addr.get("country",""),
                        "status": fields.get("status",""),
                        "severity": fields.get("severity",""),
                        "product": fields.get("product",""),
                        "product_no": fields.get("product_no",""),
                        "group": fields.get("group",""),
                        "action_plan": action_plan,
                        "hpe_last_update": last_ts,
                        "hpe_last_subject": hpe_subject,
                        "hpe_request_category": category,
                        "hpe_request_summary": summary,
                        "hpe_requested_actions": " | ".join(actions),
                        "hpe_key_links": " | ".join(key_links),
                        "event_ids": " | ".join(event_ids),
                        "problem_descriptions": " | ".join(problem_descs),
                        "ahs_links": " | ".join(ahs_links),
                        "dropbox_hosts": " | ".join(drop_hosts),
                        "dropbox_logins": " | ".join(drop_logins),
                        "comms_file": str(comms_file),
                        "generated_at": utc_now_iso()
                    }
                    results.append(row)

                    print(f"OK: {case_no} | {serial} | {fields.get('status','')}")
                except Exception as e:
                    errors.append({"case_no": case_no, "error": str(e)})
                    print(f"ERROR {case_no}: {e}")
                    try:
                        page.screenshot(path=str(outdir / "debug" / f"{case_no}_error.png"), full_page=True)
                    except Exception:
                        pass
                    try:
                        save_text(outdir / "debug" / f"{case_no}_error.html", page.content())
                    except Exception:
                        pass

            if args.format in ("csv", "both"):
                csv_path = outdir / "cases_overview.csv"
                fields_out = [
                    "case_no","serial","host_name","contact_name",
                    "addr_street","addr_city","addr_state","addr_postal_code","addr_country",
                    "status","severity","product","product_no","group","action_plan",
                    "hpe_last_update","hpe_last_subject","hpe_request_category","hpe_request_summary",
                    "hpe_requested_actions","hpe_key_links",
                    "event_ids","problem_descriptions","ahs_links","dropbox_hosts","dropbox_logins",
                    "comms_file","generated_at"
                ]
                with csv_path.open("w", encoding="utf-8", newline="") as f:
                    w = csv.DictWriter(f, fieldnames=fields_out, extrasaction="ignore")
                    w.writeheader()
                    for r in results:
                        w.writerow(r)
                print(f"\nCSV: {csv_path}")

            if args.format in ("json", "both"):
                json_path = outdir / "cases_overview.json"
                save_json(json_path, {"generated_at": utc_now_iso(), "cases": results, "errors": errors})
                print(f"JSON: {json_path}")

            if errors:
                save_json(outdir / "debug" / "errors.json", errors)
                print(f"Completed with {len(errors)} error(s). Debug in: {outdir / 'debug'}")
                return 1

            print(f"\nDone. Cases: {len(results)}")
            return 0

        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    raise SystemExit(main())