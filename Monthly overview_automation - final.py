"""VSL Performance Dashboard - Monthly Overview Report Downloader (OPTIMIZED)
===========================================================================
KEY FIXES vs previous version:
  1. Smarter chart-ready detection — polls actual Chart.js internal state
     instead of canvas pixel-sniffing (which fires too early on white frames)
  2. Single configurable WAIT_AFTER_PRINT_BTN (replaces the dual 3+3+10 s gaps)
  3. Stays on the report page between vessels — no redundant navigate-back
  4. Robust vessel-select that clears old input before typing
    5. Native browser print flow (window.print + kiosk Save as PDF)
  6. Auto-cleanup of temp Chrome profile on exit

REQUIREMENTS:
    pip install selenium webdriver-manager pandas openpyxl

USAGE:
    1. Fill in the CONFIG block below
    2. python vsl_monthly_overview_downloader.py
"""

import os, sys, time, shutil, getpass, logging, json, tempfile
from datetime import datetime

# Force UTF-8 encoding on stdout/stderr to prevent UnicodeEncodeError on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException
)
from webdriver_manager.chrome import ChromeDriverManager

# ══════════════════════════════════════════════════════════════
#  CONFIG  ← edit these
# ══════════════════════════════════════════════════════════════
VESSEL_FILE        = r"C:\Users\sonir\Downloads\Automation monthly\vessels.xlsx"
DOWNLOAD_FOLDER    = r"C:\Users\sonir\Downloads\Automation monthly\downloaded_reports\Armona"
REPORT_MONTH_LABEL = "31-May-26"          # must match exactly what appears in the table

LOGIN_URL            = "https://dashboard.vslperformance.com/#/auth/login"
MONTHLY_OVERVIEW_URL = "https://dashboard.vslperformance.com/#/main/reporting/monthly-overview-reports"

# ── Timing knobs ──────────────────────────────────────────────
#  How long to wait after clicking the site's Print button
#  before we capture the PDF.  The button triggers JS layout
#  changes (@media print reflow).  10 s is safe; lower if fast.
WAIT_AFTER_PRINT_BTN  = 13   # seconds

#  How long to wait for ALL charts to finish rendering after
#  page load + scroll.  Raise if you still see empty charts.
CHART_RENDER_TIMEOUT  = 60   # seconds 

#  Short settle pause after render is confirmed (keep it small)
CHART_SETTLE_PAUSE    = 2    # seconds

#  General Selenium element-wait timeout
SMART_TIMEOUT         = 25   # seconds

#  Pause between vessels (very short — we stay on the report page)
BETWEEN_VESSELS       = 1    # seconds
# ══════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("vsl_downloader.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────────
#  Vessel list loader
# ──────────────────────────────────────────────────────────────
def load_vessels(filepath: str) -> list:
    log.info(f"Loading vessels from: {filepath}")
    try:
        df = pd.read_csv(filepath) if filepath.lower().endswith(".csv") else pd.read_excel(filepath)
        if "vessel_name" not in df.columns:
            log.error(f"Column 'vessel_name' not found. Got: {list(df.columns)}")
            sys.exit(1)
        vessels = df["vessel_name"].dropna().str.strip().tolist()
        if not vessels:
            log.error("No vessel names found in file.")
            sys.exit(1)
        log.info(f"Loaded {len(vessels)} vessel(s): {vessels}")
        return vessels
    except FileNotFoundError:
        log.error(f"File not found: {filepath}")
        sys.exit(1)


# ──────────────────────────────────────────────────────────────
#  Chrome setup — temp profile with PDF-as-default-printer
# ──────────────────────────────────────────────────────────────
def _build_chrome_profile(download_folder: str) -> str:
    profile_dir = tempfile.mkdtemp(prefix="vsl_chrome_")
    default_dir = os.path.join(profile_dir, "Default")
    os.makedirs(default_dir, exist_ok=True)

    app_state = json.dumps({
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isHeaderFooterEnabled": False,
        "marginsType": 2,
        "isCssBackgroundEnabled": True,
        "scaling": 100,
        "scalingType": 3,
        "scalingTypePdf": 3,
        "isLandscapeEnabled": False,
        "pagesPerSheet": 1,
    })

    prefs = {
        "download": {
            "default_directory": download_folder,
            "prompt_for_download": False,
            "directory_upgrade": True,
        },
        "savefile": {
            "default_directory": download_folder,
        },
        "printing": {
            "print_preview_sticky_settings": {"appState": app_state}
        },
    }
    with open(os.path.join(default_dir, "Preferences"), "w", encoding="utf-8") as f:
        json.dump(prefs, f)

    log.info(f"Chrome profile: {profile_dir}")
    return profile_dir


def build_driver(download_folder: str) -> webdriver.Chrome:
    profile_dir = _build_chrome_profile(download_folder)
    opts = Options()
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_argument("--start-maximized")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--kiosk-printing")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts
    )
    driver.implicitly_wait(0)
    driver._vsl_profile_dir = profile_dir
    return driver


# ──────────────────────────────────────────────────────────────
#  Helpers
# ──────────────────────────────────────────────────────────────
def _js(driver, script, *args):
    return driver.execute_script(script, *args)


def safe_click(driver, el):
    _js(driver, "arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.15)
    try:
        el.click()
    except Exception:
        _js(driver, "arguments[0].click();", el)


def page_ready(driver, timeout=15):
    WebDriverWait(driver, timeout).until(
        lambda d: _js(d, "return document.readyState") == "complete"
    )


# ──────────────────────────────────────────────────────────────
#  Smart chart-ready detection
#  Strategy (in order of reliability):
#   1. No spinners / loading overlays visible
#   2. Network idle (XHR + fetch monitor)
#   3. Chart.js instances all finished animating
#   4. Canvas elements have non-trivial painted pixels
#  All four must pass 3 consecutive 1-second checks.
# ──────────────────────────────────────────────────────────────
_XHR_MONITOR_JS = """
if (!window._vslXhr) {
    window._vslXhr = {pending: 0};
    const oSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function(...a) {
        window._vslXhr.pending++;
        this.addEventListener('loadend', () => {
            window._vslXhr.pending = Math.max(0, window._vslXhr.pending - 1);
        });
        oSend.apply(this, a);
    };
    const oFetch = window.fetch;
    window.fetch = function(...a) {
        window._vslXhr.pending++;
        return oFetch.apply(this, a).finally(() => {
            window._vslXhr.pending = Math.max(0, window._vslXhr.pending - 1);
        });
    };
} else {
    window._vslXhr.pending = 0;
}
"""

_READY_CHECK_JS = """
// 1. Spinners
const spinners = document.querySelectorAll(
    'mat-spinner,mat-progress-spinner,.spinner,.loading-spinner,[class*="spinner"],[class*="loader"],.overlay'
);
for (const s of spinners) {
    const st = window.getComputedStyle(s);
    if (st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0' && s.offsetParent !== null)
        return {ok: false, reason: 'spinner'};
}

// 2. Text-based loading indicators
for (const el of document.querySelectorAll('*')) {
    if (el.childNodes.length === 1 && el.childNodes[0].nodeType === 3) {
        const t = el.textContent.trim();
        if ((t === 'Saving...' || t === 'Loading...' || t === 'Loading') &&
            window.getComputedStyle(el).display !== 'none')
            return {ok: false, reason: 'loading-text'};
    }
}

// 3. Network idle
const pending = window._vslXhr ? window._vslXhr.pending : 0;
if (pending > 0) return {ok: false, reason: 'network(' + pending + ')'};

// 4. Chart.js animation finished
if (window.Chart) {
    for (const id of Object.keys(Chart.instances || {})) {
        const c = Chart.instances[id];
        if (c && c.animating) return {ok: false, reason: 'chart-animating'};
    }
}

// 5. Canvas has painted pixels (ALL visible, non-trivial canvases must be painted)
const canvases = document.querySelectorAll('canvas');
for (const c of canvases) {
    if (c.width < 10 || c.height < 10) continue;
    const style = window.getComputedStyle(c);
    if (style.display === 'none' || style.visibility === 'hidden' || c.offsetParent === null) continue;
    
    let isPainted = false;
    try {
        const ctx = c.getContext('2d');
        if (!ctx) continue;
        const d = ctx.getImageData(0, 0, Math.min(c.width, 400), Math.min(c.height, 400)).data;
        // Look for non-white, non-transparent pixels
        for (let i = 0; i < d.length; i += 4) {
            if (d[i+3] > 10 && !(d[i] > 245 && d[i+1] > 245 && d[i+2] > 245)) {
                isPainted = true;
                break;
            }
        }
    } catch(e) { isPainted = true; }
    if (!isPainted) return {ok: false, reason: 'canvas-empty(' + (c.id || 'unnamed') + ')'};
}

return {ok: true, reason: 'ready'};
"""


def wait_for_charts(driver, timeout=CHART_RENDER_TIMEOUT):
    """Wait until all charts are truly rendered. Requires 3 consecutive OK checks."""
    log.info("  Waiting for charts to fully render...")
    _js(driver, _XHR_MONITOR_JS)

    deadline = time.time() + timeout
    streak = 0

    while time.time() < deadline:
        try:
            result = _js(driver, _READY_CHECK_JS)
            if isinstance(result, dict) and result.get("ok"):
                streak += 1
                log.info(f"  Chart check {streak}/3 ✓")
                if streak >= 3:
                    break
            else:
                reason = result.get("reason", "?") if isinstance(result, dict) else str(result)
                log.debug(f"  Not ready yet: {reason}")
                streak = 0
        except StaleElementReferenceException:
            streak = 0
        time.sleep(1)
    else:
        log.warning(f"  Chart render timed out after {timeout}s — printing anyway.")

    log.info(f"  Settling {CHART_SETTLE_PAUSE}s...")
    time.sleep(CHART_SETTLE_PAUSE)


def scroll_full_page(driver):
    """Scroll to bottom in steps (triggers lazy-render), then back to top."""
    log.info("  Scrolling page to trigger lazy sections...")
    total = _js(driver, "return document.body.scrollHeight")
    pos = 0
    while pos < total:
        pos = min(pos + 600, total)
        _js(driver, f"window.scrollTo(0, {pos});")
        time.sleep(0.3)
        total = _js(driver, "return document.body.scrollHeight")  # may grow
    _js(driver, "window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    _js(driver, "window.scrollTo(0, 0);")
    time.sleep(0.5)


# ──────────────────────────────────────────────────────────────
#  Login
# ──────────────────────────────────────────────────────────────
def login(driver, username: str, password: str):
    log.info("Logging in...")
    driver.get(LOGIN_URL)
    page_ready(driver)
    time.sleep(2)

    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='email' or @type='text']"))
    ).send_keys(username)

    pwd = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
    )
    pwd.send_keys(password)
    pwd.send_keys(Keys.RETURN)

    WebDriverWait(driver, 30).until(EC.presence_of_element_located(
        (By.XPATH, "//*[contains(text(),'MACHINERY') or contains(text(),'Fleet') or contains(text(),'DATA')]")
    ))
    log.info("  Login OK.")


# ──────────────────────────────────────────────────────────────
#  Go to Monthly Overview list
# ──────────────────────────────────────────────────────────────
def go_to_list(driver):
    log.info("  Navigating to Monthly Overview list...")
    driver.get(MONTHLY_OVERVIEW_URL)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located(
        (By.XPATH, "//*[contains(text(),'Overview') and contains(text(),'Report')]")
    ))
    # Wait for at least one enabled input (vessel search)
    WebDriverWait(driver, 20).until(
        lambda d: any(i.is_displayed() and i.is_enabled() for i in d.find_elements(By.XPATH, "//input"))
    )
    time.sleep(1)
    _js(driver, _XHR_MONITOR_JS)
    log.info("  List page ready.")


# ──────────────────────────────────────────────────────────────
#  Select vessel in the search box
# ──────────────────────────────────────────────────────────────
def select_vessel(driver, vessel_name: str):
    log.info(f"  Selecting vessel: {vessel_name}")
    inp = None
    for el in driver.find_elements(By.XPATH, "//input"):
        if el.is_displayed() and el.is_enabled():
            inp = el
            break
    if not inp:
        raise NoSuchElementException("No visible input on page.")

    safe_click(driver, inp)
    # Clear existing text robustly
    inp.send_keys(Keys.CONTROL + "a")
    inp.send_keys(Keys.DELETE)
    time.sleep(0.2)
    _js(driver, "arguments[0].value = '';", inp)
    inp.send_keys(vessel_name)
    time.sleep(0.3)

    # Accept dropdown suggestion or just press Enter
    try:
        WebDriverWait(driver, 4).until(lambda d: any(
            el.is_displayed()
            for el in d.find_elements(By.XPATH,
                f"//*[@role='option' or contains(@class,'option') or contains(@class,'item')]"
                f"[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{vessel_name[:4].upper()}')]"
            )
        ))
    except TimeoutException:
        pass  # No dropdown — Enter will still submit
    inp.send_keys(Keys.RETURN)

    # Wait for table rows to appear
    try:
        WebDriverWait(driver, SMART_TIMEOUT).until(
            lambda d: d.find_elements(By.XPATH, "//tr/td") or
                      d.find_elements(By.XPATH, "//*[contains(text(),'No Data')]")
        )
    except TimeoutException:
        log.warning("  Timed out waiting for table — continuing.")
    time.sleep(1)


# ──────────────────────────────────────────────────────────────
#  Click "View Analysis" for the target month row
# ──────────────────────────────────────────────────────────────
def click_view_analysis(driver, month_label: str):
    log.info(f"  Looking for row: {month_label}")
    try:
        row = WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//tr[td[contains(text(),'{month_label}')]]")
            )
        )
    except TimeoutException:
        dates = [el.text for el in driver.find_elements(By.XPATH, "//tr/td[1]") if el.text.strip()]
        raise RuntimeError(
            f"Month '{month_label}' not found in table.\n"
            f"Available rows: {dates}\n"
            f"Update REPORT_MONTH_LABEL to match exactly."
        )

    _js(driver, "arguments[0].scrollIntoView({block:'center'});", row)
    time.sleep(0.2)

    btn = None
    for xpath in [
        ".//button[contains(text(),'View Analysis')]",
        ".//a[contains(text(),'View Analysis')]",
        ".//button[contains(@class,'btn')]",
    ]:
        try:
            btn = row.find_element(By.XPATH, xpath)
            break
        except NoSuchElementException:
            continue
    if not btn:
        raise NoSuchElementException("'View Analysis' button not found in row.")

    safe_click(driver, btn)
    log.info("  Clicked View Analysis.")

    # Wait for Print button to appear (confirms report page loaded)
    WebDriverWait(driver, SMART_TIMEOUT).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls"))
    )
    log.info("  Report page loaded. Scrolling to trigger lazy charts...")
    time.sleep(2)  # brief settle before scroll
    scroll_full_page(driver)
    wait_for_charts(driver)


# ──────────────────────────────────────────────────────────────
#  Native print helpers (browser print engine + Save as PDF)
# ──────────────────────────────────────────────────────────────
def _pdf_snapshot(folder: str):
    snap = {}
    for name in os.listdir(folder):
        if name.lower().endswith(".pdf"):
            p = os.path.join(folder, name)
            try:
                snap[p] = (os.path.getsize(p), os.path.getmtime(p))
            except OSError:
                pass
    return snap


def _wait_new_pdf(folder: str, before: dict, timeout: int = 60):
    deadline = time.time() + timeout
    candidate = None
    stable_checks = 0

    while time.time() < deadline:
        # Ignore in-progress Chrome downloads
        if any(n.lower().endswith(".crdownload") for n in os.listdir(folder)):
            stable_checks = 0
            time.sleep(0.5)
            continue

        current = _pdf_snapshot(folder)
        new_files = [p for p in current if p not in before]
        changed_files = [p for p in current if p in before and current[p] != before[p]]
        pool = new_files + changed_files

        if pool:
            latest = max(pool, key=lambda p: os.path.getmtime(p))
            if candidate == latest:
                stable_checks += 1
            else:
                candidate = latest
                stable_checks = 1

            if stable_checks >= 3:
                # Double-check: verify if PDF has finished writing by seeking to the end
                # and checking for the standard PDF %%EOF marker.
                try:
                    if os.path.exists(candidate) and os.path.getsize(candidate) > 0:
                        with open(candidate, 'rb') as f:
                            size = os.path.getsize(candidate)
                            if size < 1024:
                                f.seek(0)
                            else:
                                f.seek(-1024, 2)
                            data = f.read()
                            if b'%%EOF' in data:
                                return candidate
                except (OSError, ValueError):
                    pass
                # Reset stable checks if it's still being written or locked
                stable_checks = 0
        else:
            stable_checks = 0

        time.sleep(0.5)

    return None


# ──────────────────────────────────────────────────────────────
#  Click Print (site flow) and save PDF produced by Chrome
# ──────────────────────────────────────────────────────────────
def capture_pdf(driver, vessel_name: str):
    log.info("  Capturing PDF via native print flow...")
    before = _pdf_snapshot(DOWNLOAD_FOLDER)

    # Enable background graphics toggle on page if present.
    try:
        bg_checkbox = driver.find_element(By.XPATH, "//div[@id='checkbox' and @role='checkbox']")
        if bg_checkbox.get_attribute("aria-checked") == "false":
            log.info("  Enabling background graphics checkbox on website...")
            safe_click(driver, bg_checkbox)
            time.sleep(1)
    except NoSuchElementException:
        pass
    except Exception as e:
        log.debug(f"  Could not interact with background graphics checkbox: {e}")

    # 1) Preferred: use page's own Print button (same as manual action).
    printed_file = None
    try:
        print_btn = WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls"))
        )
        safe_click(driver, print_btn)
        time.sleep(WAIT_AFTER_PRINT_BTN)
        printed_file = _wait_new_pdf(DOWNLOAD_FOLDER, before, timeout=30)
    except Exception as e:
        log.warning(f"  Print button flow failed, will fallback to window.print(): {e}")

    # 2) Fallback: explicitly trigger browser print.
    if not printed_file:
        log.info("  No new PDF yet; triggering window.print() fallback...")
        _js(driver, "window.print();")
        printed_file = _wait_new_pdf(DOWNLOAD_FOLDER, before, timeout=45)

    if not printed_file:
        raise RuntimeError(
            "No PDF was created by browser print flow. "
            "Check Chrome popup blocking, print permissions, and kiosk-printing compatibility."
        )

    safe_name = vessel_name.replace(" ", "").strip("")
    filename = f"{safe_name}Monthly_Overview{REPORT_MONTH_LABEL}.pdf"
    out_path = os.path.join(DOWNLOAD_FOLDER, filename)
    if os.path.exists(out_path):
        ts = datetime.now().strftime("%H%M%S")
        out_path = out_path.replace(".pdf", f"_{ts}.pdf")

    if os.path.normcase(os.path.abspath(printed_file)) != os.path.normcase(os.path.abspath(out_path)):
        # Retry rename in case Chrome still locks the file
        max_retries = 5
        for attempt in range(max_retries):
            try:
                os.replace(printed_file, out_path)
                break
            except OSError as e:
                if attempt == max_retries - 1:
                    raise
                log.warning(f"  File locked, retrying rename ({attempt + 1}/{max_retries})...")
                time.sleep(1)

    log.info(f"  [OK] Saved: {os.path.basename(out_path)}")


# ──────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────
def main():
    vessels = load_vessels(VESSEL_FILE)

    print("\n" + "=" * 60)
    print("  VSL Performance — Monthly Overview Downloader (optimized)")
    print("=" * 60)
    print(f"  Vessels : {len(vessels)}")
    print(f"  Month   : {REPORT_MONTH_LABEL}")
    print(f"  Save to : {DOWNLOAD_FOLDER}")
    print("=" * 60)

    username = input("\nUsername/email: ").strip()
    password = getpass.getpass("Password: ")

    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    driver = build_driver(DOWNLOAD_FOLDER)

    success, failed = [], []
    on_report_page = False  # True once we've clicked "View Analysis" at least once

    try:
        login(driver, username, password)
        go_to_list(driver)

        for idx, vessel in enumerate(vessels, 1):
            log.info(f"\n{'─'*55}")
            log.info(f"[{idx}/{len(vessels)}] {vessel}")
            log.info(f"{'─'*55}")
            try:
                # ── First vessel: select from list then open report
                if not on_report_page:
                    select_vessel(driver, vessel)
                    click_view_analysis(driver, REPORT_MONTH_LABEL)
                    on_report_page = True

                # ── Subsequent vessels: we're already on the report page.
                #    The vessel selector on the report page lets us switch
                #    without going back to the list.
                else:
                    log.info("  Changing vessel on report page...")
                    _js(driver, _XHR_MONITOR_JS)
                    select_vessel(driver, vessel)
                    # After selecting a new vessel the charts reload in-place
                    time.sleep(2)
                    scroll_full_page(driver)
                    wait_for_charts(driver)

                capture_pdf(driver, vessel)
                success.append(vessel)

            except Exception as exc:
                log.error(f"  [FAIL] Failed: {vessel} - {exc}")
                failed.append((vessel, str(exc)))
                on_report_page = False   # unknown state; navigate back
                try:
                    go_to_list(driver)
                except Exception:
                    pass

            time.sleep(BETWEEN_VESSELS)

    finally:
        driver.quit()
        profile = getattr(driver, "_vsl_profile_dir", None)
        if profile and os.path.exists(profile):
            shutil.rmtree(profile, ignore_errors=True)
            log.info("Cleaned up temp Chrome profile.")

    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    print(f"  Successful : {len(success)}")
    for v in success:
        print(f"      [OK] {v}")
    print(f"  Failed     : {len(failed)}")
    for v, e in failed:
        print(f"      [FAIL] {v}  ->  {e}")
    print("=" * 60)
    print(f"\nLog  : vsl_downloader.log")
    print(f"PDFs : {DOWNLOAD_FOLDER}")


if __name__ == "__main__":
    main()