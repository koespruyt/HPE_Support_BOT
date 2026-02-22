import argparse
import csv
import json
import re
import os
import time
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

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError, Error as PWError
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

# --- Global runtime deadline (optional) ---
_DEADLINE = None  # type: float | None

def _check_deadline():
    """Raise RuntimeError('GLOBAL_TIMEOUT') if a global deadline is configured and exceeded."""
    global _DEADLINE
    if _DEADLINE and time.time() > _DEADLINE:
        raise RuntimeError('GLOBAL_TIMEOUT')


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


def first_visible_locator_anywhere(page, selector_list):
    """Find first visible locator in main page OR any child frame."""
    loc = first_visible_locator(page, selector_list)
    if loc:
        return loc
    try:
        frames = list(page.frames)
        main = page.main_frame
    except Exception:
        frames, main = [], None

    for fr in frames:
        try:
            if main is not None and fr == main:
                continue
        except Exception:
            pass
        for sel in selector_list:
            try:
                l = fr.locator(sel).first
                if l.count() > 0 and l.is_visible():
                    return l
            except Exception:
                continue
    return None

def largest_visible_locator(page, selector_list):
    """Pick the *largest* visible element from a selector list.

    HPE's Cases UI often has multiple visible matches (e.g. a small header span
    containing "Cases" vs the real list container). Choosing the first match can
    accidentally scope to only one status group.
    """
    best = None
    best_area = -1.0
    for sel in selector_list:
        try:
            loc = page.locator(sel).first
            if loc.count() == 0 or not loc.is_visible():
                continue
            box = None
            try:
                box = loc.bounding_box()
            except Exception:
                box = None
            area = 0.0
            if box and box.get("width") and box.get("height"):
                area = float(box["width"]) * float(box["height"])
            if area > best_area:
                best_area = area
                best = loc
        except Exception:
            continue
    return best

def get_cases_home_url(cases_url: str) -> str:
    # Turn ".../connect/s/?tab=cases" into ".../connect/s/".
    if not cases_url:
        return "https://support.hpe.com/connect/s/"
    try:
        u = cases_url.split("?", 1)[0]
        if not u.endswith("/"):
            u += "/"
        return u
    except Exception:
        return "https://support.hpe.com/connect/s/"

def atomic_write_storage_state(context, state_path: Path) -> None:
    """Persist refreshed session state atomically.

    Many SSO sessions use *rolling* cookies/tokens. If we never write the refreshed
    state back, the on-disk state expires after ~24h and the user must re-login.
    """
    tmp = state_path.with_suffix(state_path.suffix + ".tmp")
    context.storage_state(path=str(tmp))
    tmp.replace(state_path)

def is_session_expired(page, cfg) -> bool:
    # Expired/failed-auth banners sometimes render inside IdP frames
    return any_text_present_anywhere(page, cfg.get("session_expired_text_any", []))

def log_login_status(page, cfg) -> bool:
    """Return True if we're likely authenticated (best-effort)."""
    if is_session_expired(page, cfg):
        return False
    # A couple of quick positive signals (non-exhaustive)
    if any_text_present(page, cfg.get("ready_text_any", [])):
        return True
    # When already on cases page, any case pattern is good enough.
    try:
        if page.locator("text=/\\bCase\\s+\\d{7,12}\\b/").count() > 0:
            return True
    except Exception:
        pass
    return True

def any_text_present(page, texts):
    """Return True if any of the given texts is PRESENT AND VISIBLE on the page."""
    for s in texts:
        try:
            loc = page.locator(f"text={s}").first
            if loc.count() > 0:
                try:
                    if loc.is_visible():
                        return True
                except Exception:
                    # detached / not visible
                    pass
        except Exception:
            continue
    return False

def any_text_present_anywhere(page, texts):
    """Check for any of the given texts in page OR child frames (VISIBLE only)."""
    if any_text_present(page, texts):
        return True
    try:
        frames = list(page.frames)
        main = page.main_frame
    except Exception:
        frames, main = [], None

    for fr in frames:
        try:
            if main is not None and fr == main:
                continue
        except Exception:
            pass
        for s in texts:
            try:
                loc = fr.locator(f"text={s}").first
                if loc.count() > 0:
                    try:
                        if loc.is_visible():
                            return True
                    except Exception:
                        pass
            except Exception:
                continue
    return False


def is_authenticating(page, cfg) -> bool:
    texts = cfg.get("authenticating_text_any") or []
    return any_text_present_anywhere(page, texts)


def dismiss_cookie_banner(page, cfg) -> bool:
    """Best-effort accept cookie banner (OneTrust etc.) to unblock clicks."""
    sels = cfg.get("cookie_accept_any") or []
    if not sels:
        return False
    btn = first_visible_locator_anywhere(page, sels)
    if not btn:
        return False
    try:
        btn.click(timeout=8000, force=True)
        page.wait_for_timeout(800)
        return True
    except Exception:
        return False


def wait_for_portal_state(page, cfg, timeout_ms: int = 10000) -> None:
    """Wait until the portal renders enough to decide login state.

    We avoid false positives by waiting until we can *see* at least one of:
    - a visible sign-in trigger
    - a visible sign-out trigger
    - a login/auth page (auth.hpe.com or visible login form)
    """
    start = time.time()
    signout_texts = cfg.get("signout_text_any") or ["Sign Out", "Sign out", "Log out", "Logout", "Afmelden", "Uitloggen"]
    signin_triggers = cfg.get("signin_trigger_any") or []

    while (time.time() - start) * 1000 < timeout_ms:
        try:
            if page.is_closed():
                return
        except Exception:
            return

        try:
            dismiss_cookie_banner(page, cfg)
        except Exception:
            pass

        try:
            if is_hpe_auth_page(page, cfg) or is_login_screen(page, cfg):
                return
        except Exception:
            pass

        try:
            if first_visible_locator_anywhere(page, signin_triggers):
                return
        except Exception:
            pass

        try:
            if any_text_present_anywhere(page, signout_texts):
                return
        except Exception:
            pass

        try:
            page.wait_for_timeout(250)
        except Exception:
            return



def ensure_page_alive(page, context):
    """Return a usable Page object even if the previous one was closed."""
    try:
        if page is not None and (not page.is_closed()):
            return page
    except Exception:
        pass
    try:
        if hasattr(context, "pages") and context.pages:
            # Return the newest still-open page
            for p in reversed(context.pages):
                try:
                    if not p.is_closed():
                        return p
                except Exception:
                    continue
    except Exception:
        pass
    try:
        return context.new_page()
    except Exception:
        return page

        dismiss_cookie_banner(page, cfg)

        if is_hpe_auth_page(page, cfg) or is_login_screen(page, cfg):
            return
        try:
            if is_logged_out(page, cfg):
                return
        except Exception:
            pass
        try:
            if any_text_present_anywhere(page, cfg.get("signout_text_any", [])):
                return
        except Exception:
            pass
        try:
            page.wait_for_timeout(250)
        except Exception:
            return


def is_logged_out(page, cfg) -> bool:
    """Detect logged-out state on the HPE portal home (no login form yet)."""
    # If we can see a Sign In trigger and we do NOT see a Sign Out trigger, treat as logged out.
    signout_texts = cfg.get("signout_text_any") or ["Sign Out", "Sign out", "Log out", "Logout", "Afmelden", "Uitloggen"]
    try:
        if any_text_present_anywhere(page, signout_texts):
            return False
    except Exception:
        pass

    triggers = cfg.get("signin_trigger_any") or []
    if triggers:
        if first_visible_locator_anywhere(page, triggers):
            return True

    # Structural sign-in links (often present even if the menu is closed / hidden)
    # This helps avoid false "LOGIN OK" when the SPA renders a hidden Sign In menu item.
    try:
        if page.locator("a[data-key='SignIn'], a[href*='/hpesc/public/home/signin'], a[href*='/home/signin']").count() > 0:
            return True
    except Exception:
        pass


    # Fallback: if the page contains common logged-out texts, assume logged out.
    logged_out_texts = cfg.get("logged_out_text_any") or []
    if logged_out_texts:
        try:
            if any_text_present_anywhere(page, logged_out_texts):
                return True
        except Exception:
            pass
    return False


def click_sign_in(page, context, cfg):
    """Click a visible Sign In entry/button if present.

    Some flows open a new tab/window; return the active page to continue automation.
    """
    triggers = cfg.get("signin_trigger_any") or []
    if not triggers:
        return page
    btn = first_visible_locator_anywhere(page, triggers)
    if not btn:
        return page
    try:
        try:
            btn.scroll_into_view_if_needed(timeout=2000)
        except Exception:
            pass
        new_page = None
        # Try to detect popup/new page
        try:
            with context.expect_page(timeout=3000) as pi:
                btn.click(timeout=15000, force=True)
            new_page = pi.value
        except Exception:
            # No popup; click already happened in same page.
            try:
                btn.click(timeout=15000, force=True)
            except Exception:
                pass

        try:
            page.wait_for_timeout(500)
        except Exception:
            pass

        if new_page:
            try:
                new_page.wait_for_load_state("domcontentloaded", timeout=15000)
            except Exception:
                pass
            return new_page
        else:
            try:
                page.wait_for_load_state("domcontentloaded", timeout=15000)
            except Exception:
                pass
            return page
    except Exception:
        return page


def is_hpe_auth_page(page, cfg) -> bool:
    """Detect HPE's own auth/sign-in page (auth.hpe.com) even if selectors don't match."""
    try:
        u = (page.url or "").lower()
        if "auth.hpe.com" in u:
            return True
    except Exception:
        pass
    # Fallback by visible texts (works across minor DOM changes)
    auth_texts = cfg.get("auth_page_text_any") or [
        "HPE Sign In",
        "User ID / Email Address",
        "Forgot Password",
        "Forgot User ID",
        "Unlock Account",
    ]
    try:
        return any_text_present_anywhere(page, auth_texts)
    except Exception:
        return False


def is_login_screen(page, cfg) -> bool:
    """Detect presence of a login form (supports iframe-based IdPs and HPE auth page)."""
    if is_hpe_auth_page(page, cfg):
        return True
    sels_user = cfg.get("login_username_any") or []
    sels_pass = cfg.get("login_password_any") or []
    # If either field is visible, we are on a login screen.
    if sels_user:
        if first_visible_locator_anywhere(page, sels_user):
            return True
    if sels_pass:
        if first_visible_locator_anywhere(page, sels_pass):
            return True
    # Last resort: any visible password input
    try:
        if page.locator("input[type='password']").first.is_visible(timeout=500):
            return True
    except Exception:
        pass
    return False


def perform_login(page, context, cfg, home_url: str, username: str, password: str, timeout_ms: int = 120000) -> bool:
    """More robust auto-login (no MFA).

    Strategy:
    - Always start from cfg['signin_direct_url'] (stable entry point)
    - Search inputs in page OR frames
    - Fallback to generic visible email/text/password inputs
    - Handle cookie banners + authenticating screens
    - Dump debug screenshot/html on failure (outdir/debug) using env:HPE_OUTDIR
    """
    if not username or not password:
        return False

    signin_url = (cfg.get("signin_direct_url") or "https://support.hpe.com/hpesc/public/home/signin").strip()

    def _out_debug_dir():
        outdir = os.environ.get("HPE_OUTDIR") or ""
        if not outdir:
            return None
        d = Path(outdir) / "debug"
        try:
            d.mkdir(parents=True, exist_ok=True)
        except Exception:
            return None
        return d

    def _dump_debug(tag: str):
        d = _out_debug_dir()
        if not d:
            return
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            page.screenshot(path=str(d / f"login_{tag}_{ts}.png"), full_page=True)
            save_text(d / f"login_{tag}_{ts}.html", page.content())
        except Exception:
            return

    def _first_visible_generic_user_input():
        candidates = [
            "input[type='email']",
            "input[autocomplete='username']",
            "input[name*='user' i]",
            "input[id*='user' i]",
            "input[name*='email' i]",
            "input[id*='email' i]",
            "form input[type='text']",
            "input[type='text']",
        ]
        return first_visible_locator_anywhere(page, candidates)

    def _first_visible_generic_pass_input():
        candidates = [
            "input[type='password']",
            "input[autocomplete='current-password']",
            "input[name*='pass' i]",
            "input[id*='pass' i]",
        ]
        return first_visible_locator_anywhere(page, candidates)

    def _click_first_visible(selectors):
        loc = first_visible_locator_anywhere(page, selectors)
        if not loc:
            return False
        try:
            loc.scroll_into_view_if_needed(timeout=2000)
        except Exception:
            pass
        try:
            loc.click(timeout=15000, force=True)
            return True
        except Exception:
            return False

    # selectors from cfg
    sels_user   = cfg.get("login_username_any") or []
    sels_pass   = cfg.get("login_password_any") or []
    sels_next   = cfg.get("login_next_any") or []
    sels_submit = cfg.get("login_submit_any") or []

    start = time.time()

    # Always start from stable sign-in endpoint
    try:
        page.goto(signin_url, wait_until="domcontentloaded", timeout=60000)
        dismiss_cookie_banner(page, cfg)
    except Exception:
        try:
            page.goto(home_url, wait_until="domcontentloaded", timeout=60000)
            _check_deadline()
            dismiss_cookie_banner(page, cfg)
            page.goto(signin_url, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            _dump_debug("navfail")
            return False

    while (time.time() - start) * 1000 < timeout_ms:
        page = ensure_page_alive(page, context)

        # Cookie banners can re-appear after redirects
        try:
            dismiss_cookie_banner(page, cfg)
        except Exception:
            pass

        # If we got back to portal already, consider login ok
        try:
            u = (page.url or "").lower()
            if "support.hpe.com/connect/s" in u and (not is_logged_out(page, cfg)) and (not is_login_screen(page, cfg)):
                return True
        except Exception:
            pass

        # Authenticating screen -> wait for redirect back
        try:
            if is_authenticating(page, cfg):
                try:
                    page.wait_for_url(re.compile(r".*support\.hpe\.com/connect/s/.*"), timeout=60000)
                except Exception:
                    pass
        except Exception:
            pass

        # Hard error banners
        if is_session_expired(page, cfg):
            _dump_debug("sessionexpired")
            return False

        # Find inputs (selectors first, then generic fallback)
        user_box = first_visible_locator_anywhere(page, sels_user) if sels_user else None
        if not user_box:
            user_box = _first_visible_generic_user_input()
        pass_box = first_visible_locator_anywhere(page, sels_pass) if sels_pass else None
        if not pass_box:
            pass_box = _first_visible_generic_pass_input()

        # If we have username but no password yet: fill + Next
        if user_box and (not pass_box):
            try:
                user_box.click(timeout=8000, force=True)
            except Exception:
                pass
            try:
                user_box.fill(username, timeout=15000)
            except Exception:
                try:
                    user_box.press("Control+A")
                    user_box.type(username, delay=25)
                except Exception:
                    _dump_debug("filluserfail")
                    return False

            # Next (or Enter)
            if sels_next and _click_first_visible(sels_next):
                pass
            else:
                try:
                    user_box.press("Enter")
                except Exception:
                    pass

            try:
                page.wait_for_timeout(1200)
            except Exception:
                pass
            continue

        # If we have both user+pass visible (1-step), fill user too (if empty)
        if user_box and pass_box:
            try:
                cur = (user_box.input_value(timeout=1000) or "").strip()
            except Exception:
                cur = ""
            if not cur:
                try:
                    user_box.click(timeout=8000, force=True)
                except Exception:
                    pass
                try:
                    user_box.fill(username, timeout=15000)
                except Exception:
                    try:
                        user_box.press("Control+A")
                        user_box.type(username, delay=25)
                    except Exception:
                        _dump_debug("filluserfail_1step")
                        return False

        # If we have password: fill + Submit
        if pass_box:
            try:
                pass_box.click(timeout=8000, force=True)
            except Exception:
                pass
            try:
                pass_box.fill(password, timeout=15000)
            except Exception:
                try:
                    pass_box.press("Control+A")
                    pass_box.type(password, delay=25)
                except Exception:
                    _dump_debug("fillpassfail")
                    return False

            did_click = False
            if sels_submit:
                did_click = _click_first_visible(sels_submit)
            if (not did_click) and sels_next:
                did_click = _click_first_visible(sels_next)
            if not did_click:
                try:
                    pass_box.press("Enter")
                except Exception:
                    pass

            # Wait for redirect back
            try:
                page.wait_for_url(re.compile(r".*support\.hpe\.com/connect/s/.*"), timeout=60000)
            except Exception:
                try:
                    page.wait_for_timeout(2000)
                except Exception:
                    pass

            try:
                u = (page.url or "").lower()
                if "support.hpe.com/connect/s" in u and (not is_logged_out(page, cfg)) and (not is_login_screen(page, cfg)):
                    return True
            except Exception:
                pass

            _dump_debug("postsubmit_notloggedin")
            try:
                page.wait_for_timeout(1500)
            except Exception:
                pass
            continue

        # No inputs found yet; short wait
        try:
            page.wait_for_timeout(800)
        except Exception:
            pass

    _dump_debug("timeout")
    return False



def recover_session(page, context, cfg, home_url: str, username: str, password: str) -> bool:
    """Attempt to recover from session expiration by clearing cookies/storage and logging in again."""
    page = ensure_page_alive(page, context)
    try:
        context.clear_cookies()
    except Exception:
        pass
    try:
        page.goto(home_url, wait_until="domcontentloaded", timeout=60000)
        page.evaluate("() => { try { localStorage.clear(); sessionStorage.clear(); } catch(e){} }")
    except Exception:
        pass
    ok = perform_login(page, context, cfg, home_url, username, password, timeout_ms=120000)
    return ok

def ensure_ready(page, cfg, timeout_ms=45000):
    start = datetime.now()
    while (datetime.now() - start).total_seconds() * 1000 < timeout_ms:

        # If we got redirected to auth/login, bail out early so caller can re-login.
        try:
            u = (page.url or "").lower()
            if "auth.hpe.com" in u:
                raise RuntimeError("LOGIN_REQUIRED")
        except Exception:
            pass
        if is_login_screen(page, cfg) or is_logged_out(page, cfg) or is_authenticating(page, cfg):
            raise RuntimeError("LOGIN_REQUIRED")

        # We only consider the page "ready" when we are truly on the cases tab
        try:
            if "tab=cases" not in (page.url or ""):
                page.wait_for_timeout(300)
                continue
        except Exception:
            pass

        if any_text_present_anywhere(page, cfg.get("session_expired_text_any", [])):
            raise RuntimeError("SESSION_EXPIRED")

        # Prefer: case list container or a case-number pattern
        if largest_visible_locator(page, cfg.get("case_list_container_any", [])):
            return True

        # Regex check (raw string avoids Python escape warnings)
        try:
            if page.locator("text=/\\bCase\\s+\\d{7,12}\\b/").count() > 0:
                return True
        except Exception:
            # If the target closed mid-check, let caller retry.
            raise

        page.wait_for_timeout(500)
    raise RuntimeError("CASES_PAGE_NOT_READY_TIMEOUT")


def dismiss_contract_banner(page, cfg):
    if not any_text_present(page, cfg.get("contract_banner_text_any", [])):
        return False


def goto_cases_with_retry(page, cfg, cases_url: str, timeout_ms: int = 60000) -> None:
    """Navigate to the cases page robustly.

    The HPE portal is a SPA; in some environments Playwright may surface net::ERR_ABORTED
    even though the client-side router has already switched the view. In headless runs
    this is more likely. Treat ERR_ABORTED as recoverable and verify the route.
    """
    try:
        page.goto(cases_url, wait_until="domcontentloaded", timeout=timeout_ms)
        return
    except Exception as e:
        if "net::ERR_ABORTED" not in str(e):
            raise
        print("WARN: Page.goto(cases) aborted (net::ERR_ABORTED). Retrying via UI navigation...")
        try:
            page.wait_for_timeout(1000)
        except Exception:
            pass

    # If the SPA already switched routes, we're good.
    try:
        if "tab=cases" in (page.url or ""):
            return
    except Exception:
        pass

    # Short wait for route change (if it is still in progress).
    try:
        page.wait_for_url("**tab=cases**", timeout=5000)
        return
    except Exception:
        pass

    # Click a cases link/tab in the UI.
    try:
        link = page.locator("a[href*='tab=cases']").first
        if link.count() > 0:
            try:
                link.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            link.click(timeout=15000, force=True)
            page.wait_for_url("**tab=cases**", timeout=timeout_ms)
            return
    except Exception:
        pass

    # Last resort: retry navigation with a weaker readiness condition.
    try:
        page.goto(cases_url, wait_until="commit", timeout=timeout_ms)
    except Exception as e:
        if "net::ERR_ABORTED" in str(e):
            return
        raise

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


def extract_onsite_from_comms(text: str) -> Dict[str, str]:
    """Best-effort extraction of onsite service signals from Communications.

    We keep it conservative (only fill fields when we are reasonably sure),
    because this info is mainly used to avoid UNKNOWN for In Progress cases.
    """
    out: Dict[str, str] = {}
    if not text:
        return out

    low = text.lower()
    if ("onsite" not in low) and ("on the way to your location" not in low) and ("is on the way to your location" not in low):
        return out

    out["onsite_detected"] = "1"

    # Example: "Jan Vanroy is on the way to your location..."
    m = re.search(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+is\s+on\s+the\s+way\s+to\s+your\s+location\b", text)
    if m:
        out["onsite_engineer"] = m.group(1).strip()

    # Example: "assist you with your onsite task (5401149164-541)."
    m = re.search(r"\bonsite\s+task\s*\(\s*([0-9]{7,12}-[0-9]{1,4})\s*\)", text, re.I)
    if m:
        out["onsite_task_ref"] = m.group(1).strip()

    # Some templates contain a numeric task id (different from case number)
    m = re.search(r"\bTask\s*ID\s*[:\s]+([0-9]{4,})\b", text, re.I)
    if m:
        out["onsite_task_id"] = m.group(1).strip()

    return out


def extract_onsite_kv(text: str) -> Dict[str, str]:
    """Extract key/value fields from the Onsite Service tab text."""
    out: Dict[str, str] = {}
    if not text:
        return out

    m = re.search(r"\bTask\s*ID\s+([0-9]{4,})\b", text, re.I)
    if m:
        out["onsite_task_id"] = m.group(1).strip()

    m = re.search(r"\bScheduling\s+Status\s+([A-Za-z][A-Za-z \-]{0,40})\b", text, re.I)
    if m:
        out["onsite_scheduling_status"] = m.group(1).strip()

    m = re.search(r"\bLatest\s+Service\s+Start\s+([A-Za-z]{3}\s+\d{1,2},\s+\d{4},?\s+\d{1,2}:\d{2}\s+[AP]M)\b", text)
    if m:
        out["onsite_latest_service_start"] = m.group(1).strip()

    return out


def try_extract_onsite_tab_text(page) -> str:
    """Try to open the 'Onsite Service' tab and return its visible panel text.

    Best-effort only; must never break the run.

    Note: In headless/scheduled runs the tab strip can be partially off-screen or inside an
    overflow container. We therefore scroll into view and use aria-controls when available.
    """
    markers = ["Task ID", "Scheduling Status", "Latest Service Start", "Onsite Service Request"]
    try:
        tab = page.get_by_role("tab", name=re.compile(r"Onsite\s+Service", re.I)).first
        if tab.count() == 0:
            return ""

        # Make the tab clickable even in headless/overflowed layouts
        try:
            tab.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            try:
                tab.evaluate("el => el.scrollIntoView({block:'center', inline:'center'})")
            except Exception:
                pass

        # Click the tab (force helps in overflow containers)
        try:
            tab.click(timeout=8000, force=True)
        except Exception:
            try:
                page.locator("text=/Onsite\\s+Service/i").first.click(timeout=8000, force=True)
            except Exception:
                return ""

        # Give the UI a moment to render the tabpanel content
        try:
            page.wait_for_timeout(900)
        except Exception:
            pass

        # Preferred: resolve the tabpanel via aria-controls (more deterministic than scanning all panels)
        try:
            panel_id = tab.get_attribute("aria-controls")
        except Exception:
            panel_id = None

        if panel_id:
            try:
                panel = page.locator(f"#{panel_id}")
                try:
                    panel.wait_for(state="visible", timeout=8000)
                except Exception:
                    pass
                t = panel.inner_text(timeout=8000)
                if any(m in t for m in markers):
                    return t
            except Exception:
                pass

        # Fallback: scan visible tabpanels and pick the one with the expected markers
        try:
            panels = page.get_by_role("tabpanel")
            count = panels.count()
        except Exception:
            count = 0

        for i in range(min(count, 20)):
            p = panels.nth(i)
            try:
                if not p.is_visible():
                    continue
                t = p.inner_text(timeout=8000)
                if any(m in t for m in markers):
                    return t
            except Exception:
                continue

        # Last resort: search the page main content (if the UI inlined the tab content)
        try:
            t = page.locator("main").inner_text(timeout=8000)
            return t if any(m in t for m in markers) else ""
        except Exception:
            return ""
    except Exception:
        return ""



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

    # Onsite service / dispatch updates (often for In Progress cases)
    if not actions and (("onsite service" in c) or ("onsite task" in c) or ("on the way to your location" in c) or ("is on the way to your location" in c) or ("dispatch" in c and "engineer" in c)):
        category = "ONSITE_SERVICE"
        summary = "HPE onsite interventie/dispatch loopt (technieker gepland/onderweg)."
        actions.append("Check Onsite Service tab voor planning/status (task ID, scheduling status, latest service start).")
        actions.append("Zorg dat toegang/site contact klopt; bereid interventie/onderdelen/remote access voor.")
        return category, summary, actions

    # General in-progress cases where HPE is working and no explicit customer action is detected
    if not actions and ("in progress" in st or "in progress" in c):
        category = "IN_PROGRESS"
        summary = "Case is in progress bij HPE (lopende opvolging, geen duidelijke customer action gedetecteerd)."
        actions.append("Opvolgen: check laatste HPE update; reageer enkel als HPE iets vraagt.")
        return category, summary, actions


    if not actions and "awaiting customer action" in st:
        category = "AWAITING"
        summary = "Case staat op Awaiting Customer Action."
        actions.append("Customer action vereist: check laatste HPE communicatie voor exacte vraag.")

    if not actions:
        summary = "Onvoldoende signalen om vraag automatisch te bepalen."
        actions.append("Manuele check Communications nodig.")

    return category, summary, actions

def collect_case_numbers(page, cfg, max_cases: int):
    # IMPORTANT: use the *largest* visible candidate, not the first.
    # This avoids scoping to a single status section like "Awaiting Customer Action".
    container = largest_visible_locator(page, cfg.get("case_list_container_any", []))
    scope = container if container else page
    rx = re.compile(cfg.get("case_text_regex", r"\bCase\s+(\d{7,12})\b"))
    found, seen = [], set()

    # Try to get a scrollable element handle (virtualized lists often live in a scroll panel)
    scroll_handle = None
    try:
        if container:
            scroll_handle = container.evaluate_handle(
                r"""el => {
  function isScrollable(x){
    if(!x) return false;
    const s = getComputedStyle(x);
    const oy = (s.overflowY||'').toLowerCase();
    return (oy === 'auto' || oy === 'scroll') && x.scrollHeight > x.clientHeight + 10;
  }
  let x = el;
  while (x && x !== document.body) {
    if (isScrollable(x)) return x;
    x = x.parentElement;
  }
  // Fallback: find any scrollable element that contains case numbers
  const rx = /\bCase\s+\d{7,12}\b/;
  const nodes = Array.from(document.querySelectorAll('*'));
  for (const n of nodes) {
    try {
      if (!rx.test(n.innerText || '')) continue;
      if (isScrollable(n)) return n;
    } catch (e) {}
  }
  return document.scrollingElement || document.documentElement;
}"""
            )
    except Exception:
        scroll_handle = None

    for _ in range(15):
        try:
            # Use page-wide locator to avoid missing cases outside the chosen scope.
            # Scope is still helpful for scrolling, not for matching.
            loc = page.locator("text=/\\bCase\\s+\\d{7,12}\\b/")
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
            if scroll_handle:
                scroll_handle.evaluate("el => { el.scrollTop = el.scrollTop + Math.max(800, el.clientHeight * 0.9); }")
            elif container:
                container.evaluate("el => { el.scrollTop = el.scrollTop + Math.max(800, el.clientHeight * 0.9); }")
            else:
                page.mouse.wheel(0, 1200)
        except Exception:
            try:
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
    ap.add_argument("--timeout", type=int, default=0, help="Max runtime in seconds (0=unlimited)")
    args = ap.parse_args()

    state_path = Path(args.state)
    cfg_path = Path(args.selectors)
    outdir = Path(args.outdir)
    (outdir / "cases").mkdir(parents=True, exist_ok=True)
    (outdir / "debug").mkdir(parents=True, exist_ok=True)

    # Let login/debug helpers know where to write artifacts
    os.environ.setdefault("HPE_OUTDIR", str(outdir))

    # Configure global runtime deadline (optional)
    global _DEADLINE
    _DEADLINE = (time.time() + args.timeout) if getattr(args, 'timeout', 0) and args.timeout > 0 else None

    if not state_path.exists():
        print(f"WARN: Missing state file: {state_path} (will attempt auto-login if HPE_USERNAME/HPE_PASSWORD set)")


    if not cfg_path.exists():
        print(f"ERROR: Missing selectors file: {cfg_path}")
        return 3

    cfg = load_json(cfg_path)
    cases_url = cfg.get("cases_url", "https://support.hpe.com/connect/s/?tab=cases")
    home_url = cfg.get("home_url", get_cases_home_url(cases_url))
    # Avoid emoji/unicode in console output (Windows codepages can mangle it)
    print(f"Open cases: {cases_url}")

    # Optional headless auto-login credentials (set by Run-HPECaseBot.ps1 via DPAPI)
    username = (os.environ.get("HPE_USERNAME") or "").strip()
    password = os.environ.get("HPE_PASSWORD") or ""
    if username:
        print(f"INFO: HPE_USERNAME provided (auto-login enabled): {username}")


    results, errors = [], []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(storage_state=str(state_path)) if state_path.exists() else browser.new_context()
        page = context.new_page()

        # Headless runs (Scheduled Task) sometimes render a narrower UI; use a wider viewport
        # so tabs like 'Onsite Service' remain clickable and their tabpanels load reliably.
        try:
            page.set_viewport_size({'width': 1600, 'height': 1000})
        except Exception:
            pass

        try:
            # 1) First hit the portal home to validate session quickly and log it.
            page.goto(home_url, wait_until="domcontentloaded", timeout=60000)
            _check_deadline()
            dismiss_contract_banner(page, cfg)
            dismiss_cookie_banner(page, cfg)
            # Wait a moment for SPA to render Sign In / Sign Out state to avoid false LOGIN OK.
            wait_for_portal_state(page, cfg, timeout_ms=10000)

            need_login = is_session_expired(page, cfg) or is_login_screen(page, cfg) or is_authenticating(page, cfg) or is_logged_out(page, cfg)
            if need_login:
                if username and password:
                    print("WARN: Login required. Attempting auto-login (no MFA)...")
                    ok = perform_login(page, context, cfg, home_url, username, password, timeout_ms=120000)
                    _check_deadline()
                    # Login flows may open a new tab/page; ensure we keep using the active page.
                    page = ensure_page_alive(page, context)
                    dismiss_contract_banner(page, cfg)
                    dismiss_cookie_banner(page, cfg)
                    if not ok:
                        write_alarm(outdir, args.alarm_file, "Login required but auto-login failed.", args.alarm_cmd)
                        print("LOGIN FAILED (auto-login failed; alarm written)")
                        return 2
                    # Save state as soon as we have a valid session (prevents later SESSION_EXPIRED loops)
                    try:
                        atomic_write_storage_state(context, state_path)
                        print(f"Session state refreshed: {state_path}")
                    except Exception as e:
                        print(f"WARN: could not write state file ({state_path}): {e}")
                    try:
                        page.goto(home_url, wait_until="domcontentloaded", timeout=60000)
                        _check_deadline()
                        dismiss_contract_banner(page, cfg)
                        dismiss_cookie_banner(page, cfg)
                    except Exception:
                        pass
                else:
                    write_alarm(outdir, args.alarm_file, "Session expired / login required on HPE portal.", args.alarm_cmd)
                    print("LOGIN FAILED / SESSION_EXPIRED (alarm written)")
                    return 2

            if not (is_session_expired(page, cfg) or is_login_screen(page, cfg) or is_authenticating(page, cfg) or is_logged_out(page, cfg)):
                print("LOGIN OK")


            # 2) Then go to Cases.
            goto_cases_with_retry(page, cfg, cases_url, timeout_ms=60000)
            _check_deadline()
            dismiss_contract_banner(page, cfg)

            try:
                ensure_ready(page, cfg, timeout_ms=45000)
            except (RuntimeError, PWError) as e:
                reason = str(e)
                if isinstance(e, PWError):
                    # TargetClosedError or navigation disruptions should trigger a re-login attempt.
                    reason = "LOGIN_REQUIRED"
                if reason in ("SESSION_EXPIRED","LOGIN_REQUIRED"):
                    if username and password:
                        print("WARN: SESSION_EXPIRED on cases page. Attempting auto-relogin...")
                        ok = recover_session(page, context, cfg, home_url, username, password)
                        _check_deadline()
                        page = ensure_page_alive(page, context)
                        dismiss_contract_banner(page, cfg)
                        dismiss_cookie_banner(page, cfg)
                        if not ok:
                            write_alarm(outdir, args.alarm_file, "Session expired / login required on HPE portal.", args.alarm_cmd)
                            print("SESSION_EXPIRED (auto-relogin failed; alarm written)")
                            return 2
                        try:
                            atomic_write_storage_state(context, state_path)
                            print(f"Session state refreshed: {state_path}")
                        except Exception as ex:
                            print(f"WARN: could not write state file ({state_path}): {ex}")

                        goto_cases_with_retry(page, cfg, cases_url, timeout_ms=60000)
                        _check_deadline()
                        dismiss_contract_banner(page, cfg)
                        dismiss_cookie_banner(page, cfg)
                        ensure_ready(page, cfg, timeout_ms=45000)
                    else:
                        write_alarm(outdir, args.alarm_file, "Session expired / login required on HPE portal.", args.alarm_cmd)
                        print("SESSION_EXPIRED (alarm written)")
                        return 2
                if reason != "SESSION_EXPIRED":
                    raise

            case_nos = collect_case_numbers(page, cfg, args.max)
            case_nos = case_nos[:args.max] if args.max > 0 else case_nos
            case_nos = [c for c in case_nos if c]
            print(f"Found cases: {len(case_nos)} -> " + ", ".join(case_nos[:20]) + (" ..." if len(case_nos) > 20 else ""))

            for idx, case_no in enumerate(case_nos, start=1):
                print(f"\n=== [{idx}/{len(case_nos)}] Case {case_no} ===")
                _check_deadline()
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

                    # If this looks like an onsite service case, try to enrich with a few structured fields.
                    # This ONLY adds extra JSON keys; CSV stays unchanged (extras are ignored by the writer).
                    onsite_info = extract_onsite_from_comms(comms_text)
                    onsite_hint = (
                        ("onsite" in (hpe_subject or "").lower())
                        or ("onsite_detected" in onsite_info)
                        or (category == "ONSITE_SERVICE")
                    )

                    # If comms parsing didn't reveal onsite signals, do a cheap UI hint check:
                    # (Some 'In Progress' cases have an Onsite Service tab but short/empty comms rendering.)
                    if not onsite_hint:
                        try:
                            if page.get_by_role("tab", name=re.compile(r"Onsite\s+Service", re.I)).count() > 0:
                                onsite_hint = True
                                onsite_info.setdefault("onsite_detected", "1")
                        except Exception:
                            pass

                    if onsite_hint:
                        try:
                            onsite_tab_text = try_extract_onsite_tab_text(page)
                            onsite_info.update(extract_onsite_kv(onsite_tab_text))
                        except Exception:
                            pass

                    # If onsite fields are present, prefer ONSITE_SERVICE over generic IN_PROGRESS
                    if category == "IN_PROGRESS" and any(
                        onsite_info.get(k)
                        for k in (
                            "onsite_task_id",
                            "onsite_task_ref",
                            "onsite_scheduling_status",
                            "onsite_latest_service_start",
                            "onsite_engineer",
                        )
                    ):
                        category = "ONSITE_SERVICE"
                        summary = "HPE onsite interventie/dispatch loopt (technieker gepland/onderweg)."
                        actions = [
                            "Check Onsite Service tab voor planning/status (task ID, scheduling status, latest service start).",
                            "Zorg dat toegang/site contact klopt; bereid interventie/onderdelen/remote access voor.",
                        ]

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

                        # Optional: onsite service enrichment (only in JSON; CSV ignores extras)
                        "onsite_detected": onsite_info.get("onsite_detected",""),
                        "onsite_task_ref": onsite_info.get("onsite_task_ref",""),
                        "onsite_task_id": onsite_info.get("onsite_task_id",""),
                        "onsite_scheduling_status": onsite_info.get("onsite_scheduling_status",""),
                        "onsite_latest_service_start": onsite_info.get("onsite_latest_service_start",""),
                        "onsite_engineer": onsite_info.get("onsite_engineer",""),

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

            # Refresh session state on disk to avoid daily re-login.
            try:
                atomic_write_storage_state(context, state_path)
                print(f"Session state refreshed: {state_path}")
            except Exception as e:
                print(f"WARN: could not refresh state file ({state_path}): {e}")

            if errors:
                save_json(outdir / "debug" / "errors.json", errors)
                print(f"Completed with {len(errors)} error(s). Debug in: {outdir / 'debug'}")
                return 1

            print(f"\nDone. Cases: {len(results)}")
            return 0
        except KeyboardInterrupt:
            # Graceful Ctrl+C
            print('INTERRUPTED')
            return 130

        except RuntimeError as e:
            if str(e) == 'GLOBAL_TIMEOUT':
                print('ERROR: GLOBAL_TIMEOUT reached. Aborting run.')
                try:
                    page.screenshot(path=str(outdir / 'debug' / 'global_timeout.png'), full_page=True)
                except Exception:
                    pass
                try:
                    save_text(outdir / 'debug' / 'global_timeout.html', page.content())
                except Exception:
                    pass
                return 4
            raise

        finally:
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass

if __name__ == "__main__":
    raise SystemExit(main())
