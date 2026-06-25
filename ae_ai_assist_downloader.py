"""
VSL Performance Dashboard - AE AI Assist Report Downloader

Logic:
  1. Load vessel list from Excel/CSV
  2. Login to VSL dashboard
  3. For each vessel:
     a. Navigate to AE AI Assist report page
     b. Search vessel name → Enter
     c. Find all rows matching TARGET_MONTH (e.g. "May")
     d. For each engine row (AE1, AE2, AE3...):
        - Click the View button
        - Wait for report to render
        - Click Print → download PDF
        - Save to vessel-specific subfolder as AE_<N>_May_2026.pdf
  4. Summary printed at end

REQUIREMENTS:
    pip install selenium webdriver-manager pandas openpyxl

USAGE:
    1. Fill in the CONFIG block below
    2. python ae_ai_assist_downloader.py
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
# CONFIG  ← edit these
# ══════════════════════════════════════════════════════════════

VESSEL_FILE     = r"C:\Users\sonir\Downloads\Automation AE\vessels.xlsx"
DOWNLOAD_FOLDER = r"C:\Users\sonir\Downloads\Automation AE"

# Month to download — must match the partial date text in the table (e.g. "-May-")
TARGET_MONTH    = "May"          # change to "Jun", "Apr", etc. as needed
REPORT_YEAR     = "2026"         # used in the output PDF filename

LOGIN_URL    = "https://dashboard.vslperformance.com/#/auth/login"
AE_ASSIST_URL = "https://dashboard.vslperformance.com/#/main/reporting/ae-performance-assistant"

# ── Timing knobs ──────────────────────────────────────────────
WAIT_AFTER_PRINT_BTN  = 10   # seconds — wait after clicking Print
CHART_RENDER_TIMEOUT  = 60   # seconds — max wait for charts to fully render
CHART_SETTLE_PAUSE    = 2    # seconds — brief settle after render confirmed
SMART_TIMEOUT         = 25   # seconds — general Selenium element-wait timeout
BETWEEN_ENGINES       = 2    # seconds — pause between engine downloads
BETWEEN_VESSELS       = 2    # seconds — pause between vessels

# ══════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("ae_downloader.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────────
# Vessel list loader
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
# Chrome setup — temp profile with PDF-as-default-printer
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
# Helpers
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
# Smart chart-ready detection
# Strategy (in order of reliability):
#   1. No spinners / loading overlays visible
#   2. Network idle (XHR + fetch monitor)
#   3. Chart.js instances all finished animating
#   4. Canvas elements have non-trivial painted pixels
# All four must pass 3 consecutive 1-second checks.
# ──────────────────────────────────────────────────────────────

_XHR_MONITOR_JS = """
if (!window.__vslXhr) {
    window.__vslXhr = {pending: 0};
    const oSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function(...a) {
        window.__vslXhr.pending++;
        this.addEventListener('loadend', () => {
            window.__vslXhr.pending = Math.max(0, window.__vslXhr.pending - 1);
        });
        oSend.apply(this, a);
    };
    const oFetch = window.fetch;
    window.fetch = function(...a) {
        window.__vslXhr.pending++;
        return oFetch.apply(this, a).finally(() => {
            window.__vslXhr.pending = Math.max(0, window.__vslXhr.pending - 1);
        });
    };
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
const pending = window.__vslXhr ? window.__vslXhr.pending : 0;
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
# Login
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
# Navigate to AE AI Assist list page
# ──────────────────────────────────────────────────────────────

def go_to_ae_list(driver):
    log.info("  Navigating to AE AI Assist page...")
    driver.get(AE_ASSIST_URL)
    # Wait for at least one enabled input (vessel search)
    WebDriverWait(driver, 30).until(
        lambda d: any(i.is_displayed() and i.is_enabled() for i in d.find_elements(By.XPATH, "//input"))
    )
    time.sleep(1)
    _js(driver, _XHR_MONITOR_JS)
    log.info("  AE page ready.")

# ──────────────────────────────────────────────────────────────
# Search for vessel
# ──────────────────────────────────────────────────────────────

def search_vessel(driver, vessel_name: str):
    log.info(f"  Searching vessel: {vessel_name}")
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
# Get all AE rows for target month
# ──────────────────────────────────────────────────────────────

def get_month_rows(driver, month_name: str) -> list:
    """Return all <tr> elements whose first date cell contains '-<month_name>-'."""
    rows = driver.find_elements(
        By.XPATH,
        f"//tbody/tr[td[contains(text(),'-{month_name}-')]]"
    )
    log.info(f"  Found {len(rows)} row(s) for month '{month_name}'")
    return rows

# ──────────────────────────────────────────────────────────────
# Create vessel-specific subfolder
# ──────────────────────────────────────────────────────────────

def create_vessel_folder(base_folder: str, vessel_name: str) -> str:
    safe = vessel_name.replace("/", "-").replace("\\", "-").strip()
    folder = os.path.join(base_folder, safe)
    os.makedirs(folder, exist_ok=True)
    return folder

# ──────────────────────────────────────────────────────────────
# Open individual AE report by month + engine number
# ──────────────────────────────────────────────────────────────

def open_ae_report(driver, month_name: str, engine_no: str):
    """Click the View button on the row matching month + engine number."""
    log.info(f"  Opening AE report: month={month_name}, engine={engine_no}")

    # Find the specific row: date contains -<month>- AND 2nd cell matches engine_no
    try:
        row = WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.presence_of_element_located((
                By.XPATH,
                f"//tbody/tr[td[contains(text(),'-{month_name}-')] and td[2][normalize-space()='{engine_no}']]"
            ))
        )
    except TimeoutException:
        # Fallback: try position-based match (first row with month that hasn't been clicked)
        rows = driver.find_elements(
            By.XPATH,
            f"//tbody/tr[td[contains(text(),'-{month_name}-')]]"
        )
        if not rows:
            raise RuntimeError(f"No rows found for month '{month_name}'")
        # Find row whose engine column matches
        row = None
        for r in rows:
            try:
                cell = r.find_element(By.XPATH, "./td[2]")
                if cell.text.strip() == str(engine_no):
                    row = r
                    break
            except NoSuchElementException:
                continue
        if not row:
            raise RuntimeError(f"Row for month='{month_name}' engine='{engine_no}' not found.")

    _js(driver, "arguments[0].scrollIntoView({block:'center'});", row)
    time.sleep(0.3)

    # Find and click the View button in this row
    btn = None
    for xpath in [
        ".//span[normalize-space()='View']",
        ".//button[contains(text(),'View')]",
        ".//a[contains(text(),'View')]",
        ".//button[contains(@class,'btn')]",
    ]:
        try:
            btn = row.find_element(By.XPATH, xpath)
            break
        except NoSuchElementException:
            continue
    if not btn:
        raise NoSuchElementException(f"'View' button not found in row for engine {engine_no}.")

    safe_click(driver, btn)
    log.info(f"  Clicked View for engine {engine_no}.")

    # Wait for report page to load (Print button appears)
    try:
        WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Print']"))
        )
    except TimeoutException:
        # Fallback: wait for any print-like button
        WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls"))
        )

    log.info("  Report page loaded. Scrolling to trigger lazy charts...")
    time.sleep(2)
    scroll_full_page(driver)
    wait_for_charts(driver)

# ──────────────────────────────────────────────────────────────
# Native print helpers (browser print engine + Save as PDF)
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
                return candidate
        else:
            stable_checks = 0

        time.sleep(0.5)

    return None

# ──────────────────────────────────────────────────────────────
# Click Print and capture PDF
# ──────────────────────────────────────────────────────────────

def capture_pdf(driver, vessel_name: str, engine_no: str, vessel_folder: str):
    log.info(f"  Capturing PDF for engine {engine_no}...")
    before = _pdf_snapshot(DOWNLOAD_FOLDER)

    # Enable background graphics toggle if present
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

    # 1) Preferred: use page's own Print button
    printed_file = None
    try:
        print_btn = WebDriverWait(driver, SMART_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Print']"))
        )
        safe_click(driver, print_btn)
        time.sleep(WAIT_AFTER_PRINT_BTN)
        printed_file = _wait_new_pdf(DOWNLOAD_FOLDER, before, timeout=30)
    except Exception as e:
        log.warning(f"  Print button (span) flow failed, trying button.btncls: {e}")

    # Fallback: button.btncls
    if not printed_file:
        try:
            print_btn = WebDriverWait(driver, SMART_TIMEOUT).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls"))
            )
            safe_click(driver, print_btn)
            time.sleep(WAIT_AFTER_PRINT_BTN)
            printed_file = _wait_new_pdf(DOWNLOAD_FOLDER, before, timeout=30)
        except Exception as e:
            log.warning(f"  button.btncls flow failed, will fallback to window.print(): {e}")

    # 2) Fallback: explicitly trigger browser print
    if not printed_file:
        log.info("  No new PDF yet; triggering window.print() fallback...")
        _js(driver, "window.print();")
        printed_file = _wait_new_pdf(DOWNLOAD_FOLDER, before, timeout=45)

    if not printed_file:
        raise RuntimeError(
            "No PDF was created by browser print flow. "
            "Check Chrome popup blocking, print permissions, and kiosk-printing compatibility."
        )

    # Rename and move into vessel subfolder
    safe_name = vessel_name.replace(" ", "_").strip("_")
    filename = f"{safe_name}_AE_{engine_no}_{TARGET_MONTH}_{REPORT_YEAR}.pdf"
    out_path = os.path.join(vessel_folder, filename)
    if os.path.exists(out_path):
        ts = datetime.now().strftime("%H%M%S")
        out_path = out_path.replace(".pdf", f"_{ts}.pdf")

    if os.path.normcase(os.path.abspath(printed_file)) != os.path.normcase(os.path.abspath(out_path)):
        os.replace(printed_file, out_path)

    log.info(f"  [OK] Saved: {os.path.basename(out_path)}")

# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────

def main():
    vessels = load_vessels(VESSEL_FILE)

    print("\n" + "=" * 60)
    print("  VSL Performance — AE AI Assist Downloader")
    print("=" * 60)
    print(f"  Vessels : {len(vessels)}")
    print(f"  Month   : {TARGET_MONTH} {REPORT_YEAR}")
    print(f"  Save to : {DOWNLOAD_FOLDER}")
    print("=" * 60)

    username = input("\nUsername/email: ").strip()
    password = getpass.getpass("Password: ")

    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    driver = build_driver(DOWNLOAD_FOLDER)

    success = []   # list of (vessel, engine) tuples
    failed  = []   # list of (vessel, engine, error) tuples

    try:
        login(driver, username, password)

        for idx, vessel in enumerate(vessels, 1):
            log.info(f"\n{'─'*55}")
            log.info(f"[{idx}/{len(vessels)}] {vessel}")
            log.info(f"{'─'*55}")

            vessel_folder = create_vessel_folder(DOWNLOAD_FOLDER, vessel)

            try:
                # Navigate to AE list and search vessel
                go_to_ae_list(driver)
                search_vessel(driver, vessel)

                # Find all rows matching target month
                rows = get_month_rows(driver, TARGET_MONTH)

                if not rows:
                    log.warning(f"  No rows found for month '{TARGET_MONTH}' — skipping vessel.")
                    failed.append((vessel, "N/A", f"No rows for month '{TARGET_MONTH}'"))
                    continue

                # Collect engine numbers from column 2
                engine_numbers = []
                for row in rows:
                    try:
                        eng = row.find_element(By.XPATH, "./td[2]").text.strip()
                        if eng:
                            engine_numbers.append(eng)
                    except NoSuchElementException:
                        pass

                if not engine_numbers:
                    log.warning("  Could not read engine numbers from rows.")
                    engine_numbers = [str(i) for i in range(1, len(rows) + 1)]

                log.info(f"  Engines to download: {engine_numbers}")

                for eng in engine_numbers:
                    try:
                        # Re-navigate and search for each engine (ensures fresh page state)
                        go_to_ae_list(driver)
                        search_vessel(driver, vessel)

                        open_ae_report(driver, TARGET_MONTH, eng)
                        capture_pdf(driver, vessel, eng, vessel_folder)
                        success.append((vessel, eng))

                    except Exception as exc:
                        log.error(f"  [FAIL] Engine {eng} failed: {exc}")
                        failed.append((vessel, eng, str(exc)))
                        # Navigate back to list to recover
                        try:
                            go_to_ae_list(driver)
                        except Exception:
                            pass

                    time.sleep(BETWEEN_ENGINES)

            except Exception as exc:
                log.error(f"  [FAIL] Vessel failed: {vessel} - {exc}")
                failed.append((vessel, "ALL", str(exc)))
                try:
                    go_to_ae_list(driver)
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
    for v, e in success:
        print(f"      [OK]   {v}  —  Engine {e}")
    print(f"  Failed     : {len(failed)}")
    for v, e, err in failed:
        print(f"      [FAIL] {v}  Engine {e}  ->  {err}")
    print("=" * 60)
    print(f"\nLog  : ae_downloader.log")
    print(f"PDFs : {DOWNLOAD_FOLDER}")


if __name__ == "__main__":
    main()
