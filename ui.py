import os
import sys
import json
import time
import tempfile
import shutil
import logging
import getpass
from datetime import datetime
import pandas as pd
import streamlit as st


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

# Force UTF-8 log stream
if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")

# Streamlit UI Configuration
st.set_page_config(page_title="VSL Report Downloader Dashboard", page_icon="🚢", layout="centered")

st.title("🚢 VSL Performance Report Downloader")
st.markdown("Automate report downloads for Monthly Overviews and 4-Stroke analytics engines seamlessly.")

# ─── STREAMLIT UI INPUT COLUMNS ───
st.header("1. Authentication & Credentials")
col1, col2 = st.columns(2)
with col1:
    username = st.text_input("Username / Email", placeholder="user@vslperformance.com")
with col2:
    password = st.text_input("Password", type="password", placeholder="••••••••")

st.header("2. Report Targets & Configuration")
col3, col4, col5 = st.columns(3)
with col3:
    downloader_mode = st.selectbox("Select Automation Engine", ["Monthly Overview", "4-Stroke"])
with col4:
    # Generates a sequence of recent years
    selected_year = st.selectbox("Select Year", ["2024", "2025", "2026", "2027", "2028"], index=2)
with col5:
    selected_month = st.selectbox("Select Month", ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], index=4)

# Format the REPORT_MONTH_LABEL dynamically (e.g., "May-26")
short_year = selected_year[-2:]
computed_month_label = f"{selected_month}-{short_year}"

st.markdown(f"**Target Report Label constructed:** `{computed_month_label}`")

col_folder, col_file = st.columns(2)
with col_folder:
    custom_download_path = st.text_input("Local Download Folder Path", value=r"C:\Users\sonir\Downloads\Automation monthly\downloaded_reports")
with col_file:
    uploaded_file = st.file_uploader("Attach Vessels Excel File (.xlsx)", type=["xlsx", "xls"])

# ─── CORE SELENIUM BACKEND BACKBONE ───
WAIT_AFTER_PRINT_BTN = 13
CHART_RENDER_TIMEOUT = 60
CHART_SETTLE_PAUSE = 2
SMART_TIMEOUT = 25

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
const spinners = document.querySelectorAll('mat-spinner,mat-progress-spinner,.spinner,.loading-spinner,[class*="spinner"],[class*="loader"],.overlay');
for (const s of spinners) {
    const st = window.getComputedStyle(s);
    if (st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0' && s.offsetParent !== null) return {ok: false, reason: 'spinner'};
}
for (const el of document.querySelectorAll('*')) {
    if (el.childNodes.length === 1 && el.childNodes[0].nodeType === 3) {
        const t = el.textContent.trim();
        if ((t === 'Saving...' || t === 'Loading...' || t === 'Loading') && window.getComputedStyle(el).display !== 'none') return {ok: false, reason: 'loading-text'};
    }
}
const pending = window._vslXhr ? window._vslXhr.pending : 0;
if (pending > 0) return {ok: false, reason: 'network(' + pending + ')'};
if (window.Chart) {
    for (const id of Object.keys(Chart.instances || {})) {
        const c = Chart.instances[id];
        if (c && c.animating) return {ok: false, reason: 'chart-animating'};
    }
}
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

def _build_chrome_profile(download_folder):
    profile_dir = tempfile.mkdtemp(prefix="vsl_chrome_")
    default_dir = os.path.join(profile_dir, "Default")
    os.makedirs(default_dir, exist_ok=True)
    app_state = json.dumps({
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2, "isHeaderFooterEnabled": False, "marginsType": 2, "isCssBackgroundEnabled": True,
        "scaling": 100, "scalingType": 3, "scalingTypePdf": 3, "isLandscapeEnabled": False, "pagesPerSheet": 1
    })
    prefs = {
        "download.default_directory": os.path.abspath(download_folder),
        "download.prompt_for_download": False, "download.directory_upgrade": True,
        "savefile.default_directory": os.path.abspath(download_folder),
        "printing.print_preview_sticky_settings": {"appState": app_state}
    }
    with open(os.path.join(default_dir, "Preferences"), "w", encoding="utf-8") as f:
        json.dump(prefs, f)
    return profile_dir

def build_driver(download_folder):
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
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    driver.implicitly_wait(0)
    driver._vsl_profile_dir = profile_dir
    return driver

def _js(driver, script, *args): return driver.execute_script(script, *args)

def safe_click(driver, el):
    _js(driver, "arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.15)
    try: el.click()
    except Exception: _js(driver, "arguments[0].click();", el)

def page_ready(driver, timeout=15):
    WebDriverWait(driver, timeout).until(lambda d: _js(d, "return document.readyState") == "complete")

def wait_for_charts(driver, timeout=CHART_RENDER_TIMEOUT):
    _js(driver, _XHR_MONITOR_JS)
    deadline = time.time() + timeout
    streak = 0
    while time.time() < deadline:
        try:
            result = _js(driver, _READY_CHECK_JS)
            if isinstance(result, dict) and result.get("ok"):
                streak += 1
                if streak >= 3: break
            else: streak = 0
        except StaleElementReferenceException: streak = 0
        time.sleep(1)
    time.sleep(CHART_SETTLE_PAUSE)

def scroll_full_page(driver):
    total = _js(driver, "return document.body.scrollHeight")
    pos = 0
    while pos < total:
        pos = min(pos + 600, total)
        _js(driver, f"window.scrollTo(0, {pos});")
        time.sleep(0.3)
        total = _js(driver, "return document.body.scrollHeight")
    _js(driver, "window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    _js(driver, "window.scrollTo(0, 0);")
    time.sleep(0.5)

def login(driver, username, password):
    driver.get("https://dashboard.vslperformance.com/#/auth/login")
    page_ready(driver)
    time.sleep(2)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='email' or @type='text']"))).send_keys(username)
    pwd = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
    pwd.send_keys(password)
    pwd.send_keys(Keys.RETURN)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'MACHINERY') or contains(text(),'Fleet') or contains(text(),'DATA')]")))

def go_to_list(driver, url):
    driver.get(url)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Overview') or contains(text(),'Report') or contains(text(),'four-stroke')]")))
    try:
        WebDriverWait(driver, 20).until(lambda d: any(i.is_displayed() and i.is_enabled() for i in d.find_elements(By.XPATH, "//input")))
    except:
        pass
    time.sleep(1)
    _js(driver, _XHR_MONITOR_JS)

def select_vessel(driver, vessel_name):
    inp = None
    for el in driver.find_elements(By.XPATH, "//input"):
        if el.is_displayed() and el.is_enabled():
            inp = el
            break
    if not inp: raise NoSuchElementException("No visible input on page.")
    safe_click(driver, inp)
    inp.send_keys(Keys.CONTROL + "a")
    inp.send_keys(Keys.DELETE)
    time.sleep(0.2)
    _js(driver, "arguments[0].value = '';", inp)
    inp.send_keys(vessel_name)
    time.sleep(0.3)
    try:
        WebDriverWait(driver, 4).until(lambda d: any(el.is_displayed() for el in d.find_elements(By.XPATH, f"//*[@role='option' or contains(@class,'option') or contains(@class,'item')][contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{vessel_name[:4].upper()}')]")))
    except TimeoutException: pass
    inp.send_keys(Keys.RETURN)
    try:
        WebDriverWait(driver, SMART_TIMEOUT).until(lambda d: d.find_elements(By.XPATH, "//tr/td") or d.find_elements(By.XPATH, "//*[contains(text(),'No Data')]"))
    except TimeoutException: pass
    time.sleep(1)

def click_view_analysis(driver, month_label):
    try:
        row = WebDriverWait(driver, SMART_TIMEOUT).until(EC.presence_of_element_located((By.XPATH, f"//tr[td[contains(text(),'{month_label}')]]")))
    except TimeoutException:
        raise RuntimeError(f"Month '{month_label}' not found in table.")
    _js(driver, "arguments[0].scrollIntoView({block:'center'});", row)
    time.sleep(0.2)
    btn = None
    for xpath in [".//button[contains(text(),'View Analysis')]", ".//a[contains(text(),'View Analysis')]", ".//button[contains(@class,'btn')]"]:
        try:
            btn = row.find_element(By.XPATH, xpath)
            break
        except NoSuchElementException: continue
    if not btn: raise NoSuchElementException("'View Analysis' button not found in row.")
    safe_click(driver, btn)
    WebDriverWait(driver, SMART_TIMEOUT).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls")))
    time.sleep(2)
    scroll_full_page(driver)
    wait_for_charts(driver)

def _pdf_snapshot(folder):
    snap = {}
    for name in os.listdir(folder):
        if name.lower().endswith(".pdf"):
            p = os.path.join(folder, name)
            try: snap[p] = (os.path.getsize(p), os.path.getmtime(p))
            except OSError: pass
    return snap

def _wait_new_pdf(folder, before, timeout=60):
    deadline = time.time() + timeout
    candidate = None
    stable_checks = 0
    while time.time() < deadline:
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
            if candidate == latest: stable_checks += 1
            else:
                candidate = latest
                stable_checks = 1
            if stable_checks >= 3:
                try:
                    if os.path.exists(candidate) and os.path.getsize(candidate) > 0:
                        with open(candidate, 'rb') as f:
                            size = os.path.getsize(candidate)
                            if size < 1024: f.seek(0)
                            else: f.seek(-1024, 2)
                            if b'%%EOF' in f.read(): return candidate
                except (OSError, ValueError): pass
                stable_checks = 0
        else: stable_checks = 0
        time.sleep(0.5)
    return None

def capture_pdf(driver, vessel_name, download_folder, month_label):
    before = _pdf_snapshot(download_folder)
    try:
        bg_checkbox = driver.find_element(By.XPATH, "//div[@id='checkbox' and @role='checkbox']")
        if bg_checkbox.get_attribute("aria-checked") == "false":
            safe_click(driver, bg_checkbox)
            time.sleep(1)
    except Exception: pass
    printed_file = None
    try:
        print_btn = WebDriverWait(driver, SMART_TIMEOUT).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btncls")))
        safe_click(driver, print_btn)
        time.sleep(WAIT_AFTER_PRINT_BTN)
        printed_file = _wait_new_pdf(download_folder, before, timeout=30)
    except Exception: pass
    if not printed_file:
        _js(driver, "window.print();")
        printed_file = _wait_new_pdf(download_folder, before, timeout=45)
    if not printed_file: raise RuntimeError("No PDF was created by browser print flow.")
    safe_name = vessel_name.replace(" ", "").strip()
    filename = f"{safe_name}Monthly_Overview_{month_label}.pdf"
    out_path = os.path.join(download_folder, filename)
    if os.path.exists(out_path):
        out_path = out_path.replace(".pdf", f"_{datetime.now().strftime('%H%M%S')}.pdf")
    if os.path.normcase(os.path.abspath(printed_file)) != os.path.normcase(os.path.abspath(out_path)):
        for attempt in range(5):
            try:
                os.replace(printed_file, out_path)
                break
            except OSError:
                if attempt == 4: raise
                time.sleep(1)

def run_downloader(vessel_list, url, target_month, out_dir, status_placeholder, progress_placeholder):
    out_dir = out_dir.strip('\'"')
    os.makedirs(out_dir, exist_ok=True)
    driver = build_driver(out_dir)
    
    try:
        status_placeholder.info("Initiating login session framework...")
        login(driver, username, password)
        
        status_placeholder.info("Navigating to target reports list...")
        go_to_list(driver, url)
        
        on_report_page = False
        
        for idx, vessel in enumerate(vessel_list):
            status_placeholder.markdown(f"**Processing Engine Entry [{idx+1}/{len(vessel_list)}]:** `{vessel}`")
            progress_placeholder.progress((idx + 1) / len(vessel_list))
            
            try:
                if not on_report_page:
                    select_vessel(driver, vessel)
                    click_view_analysis(driver, target_month)
                    on_report_page = True
                else:
                    _js(driver, _XHR_MONITOR_JS)
                    select_vessel(driver, vessel)
                    time.sleep(2)
                    scroll_full_page(driver)
                    wait_for_charts(driver)
                
                capture_pdf(driver, vessel, out_dir, target_month)
                status_placeholder.success(f"Downloaded {vessel} successfully.")
            except Exception as e:
                status_placeholder.error(f"Error processing {vessel}: {e}")
                on_report_page = False
                try: go_to_list(driver, url)
                except Exception: pass
            
            time.sleep(1) 
            
        status_placeholder.success("🎉 Automation run finished entirely!")
    except Exception as e:
        status_placeholder.error(f"Automation execution broken: {e}")
    finally:
        driver.quit()
        if hasattr(driver, "_vsl_profile_dir") and os.path.exists(driver._vsl_profile_dir):
            shutil.rmtree(driver._vsl_profile_dir, ignore_errors=True)

# ─── TRIGGER AUTOMATION ACTION ───
st.header("3. Run Script Engine")
if st.button("🚀 Start Report Downloader Pipeline", use_container_width=True):
    if not username or not password:
        st.error("Please insert valid profile authentication fields first.")
    elif uploaded_file is None:
        st.error("Please drop your layout vessels Excel document tracking reference table.")
    else:
        try:
            # Parse vessels right from standard attachment
            df = pd.read_excel(uploaded_file)
            if "vessel_name" not in df.columns:
                st.error("Missing explicitly targeted standard column header identifier name: 'vessel_name'")
            else:
                vessels = df["vessel_name"].dropna().str.strip().tolist()
                st.success(f"Discovered ({len(vessels)}) parsing target targets inside worksheet layout structure matrix.")
                
                # Dynamic targeted URL routing based on active drop-down execution choices
                target_url = (
                    "https://dashboard.vslperformance.com/#/main/performance/me-performance-feedback/four-stroke"
                    if downloader_mode == "4-Stroke" 
                    else "https://dashboard.vslperformance.com/#/main/performance/me-performance-feedback/monthly-overview"
                )
                
                # ─── PASTE THE NEW BLOCK HERE ───
                import threading
                from streamlit.runtime.scriptrunner import add_script_run_ctx

                # 1. Create the UI placeholders on the MAIN thread
                status_box = st.empty()
                progress_bar = st.progress(0)

                with st.spinner("Executing sequence profile matrix configurations..."):
                    # 2. Create the worker thread targeting your downloader function
                    download_thread = threading.Thread(
                        target=run_downloader, 
                        args=(vessels, target_url, computed_month_label, custom_download_path, status_box, progress_bar)
                    )
                    
                    # 3. CRITICAL: Inject the active Streamlit context into the thread
                    add_script_run_ctx(download_thread)
                    
                    # 4. Fire off the thread background sequence safely
                    download_thread.start()
                    
                    # 5. Keep the UI active and responsive while the thread runs
                    while download_thread.is_alive():
                        time.sleep(0.5)
                        
        except Exception as file_err:
            st.error(f"Could not cleanly handle reading target attachment layout config: {file_err}")