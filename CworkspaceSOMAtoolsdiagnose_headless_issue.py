"""Diagnostic tool to reproduce and analyze headless Chrome issue."""
import os, sys, time, platform, subprocess
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

def print_section(title, char="="):
    print(f"\n{char * 80}\n  {title}\n{char * 80}\n")

def check_chrome_installation():
    print_section("1. Chrome Installation Check")
    from soma_app.infra.webdriver_factory import _find_chrome_executable
    try:
        chrome_exe = _find_chrome_executable()
        print(f"[OK] Chrome found at: {chrome_exe}")
        result = subprocess.run([chrome_exe, '--version'], capture_output=True, text=True, timeout=5)
        version = result.stdout.strip() or result.stderr.strip()
        print(f"[OK] Chrome version: {version}")
        return True
    except Exception as e:
        print(f"[ERROR] Chrome not found: {e}")
        return False

def check_headless_resolution():
    print_section("2. Headless Resolution Logic Test")
    os.environ['HEADLESS'] = 'false'
    from soma_app.infra.webdriver_factory import _resolve_headless
    result1 = _resolve_headless(headless=False)
    print(f"_resolve_headless(headless=False) = {result1} {'[OK]' if not result1 else '[FAIL]'}")
    result2 = _resolve_headless(headless=True)
    print(f"_resolve_headless(headless=True) = {result2} {'[OK]' if result2 else '[FAIL]'}")
    result3 = _resolve_headless()
    print(f"_resolve_headless() with HEADLESS=false = {result3} {'[OK]' if not result3 else '[FAIL]'}")
    return all([not result1, result2, not result3])

def check_visible_chrome_launch():
    print_section("3. Visible Chrome Launch Test")
    from soma_app.infra.webdriver_factory import _launch_visible_chrome
    try:
        print("Launching Chrome with debugging port...")
        start = time.time()
        port, profile_dir, proc = _launch_visible_chrome()
        elapsed = time.time() - start
        print(f"[OK] Chrome launched in {elapsed:.1f}s (PID: {proc.pid}, Port: {port})")
        if proc.poll() is None:
            print(f"[OK] Chrome process is running")
            time.sleep(1)
            proc.terminate()
            proc.wait(timeout=5)
            print(f"[OK] Chrome terminated successfully")
            return True
        return False
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def check_selenium_attachment():
    print_section("4. Selenium Driver Attachment Test")
    os.environ.update({'GOOGLE_CREDENTIALS_PATH': 'deploy/credenciais.json', 'SPREADSHEET_URL': 'https://test', 'SITE_USER': 'test', 'SITE_PASSWORD': 'test'})
    from soma_app.infra.webdriver_factory import create_driver
    try:
        print("Creating Selenium driver with headless=False...")
        start = time.time()
        driver = create_driver(headless=False)
        elapsed = time.time() - start
        print(f"[OK] Driver created in {elapsed:.1f}s")
        print(f"  Window handles: {len(driver.window_handles)}, URL: {driver.current_url}")
        rect = driver.get_window_rect()
        print(f"  Window size: {rect['width']}x{rect['height']}")
        driver.get("https://www.example.com")
        time.sleep(1)
        if driver.title:
            print(f"[OK] Navigation successful, Title: {driver.title}")
        driver.quit()
        print(f"[OK] Driver quit successfully")
        return True
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def main():
    print("\n" + "="*80 + f"\n  SOMA Headless Chrome Diagnostic\n  Platform: {platform.platform()}\n" + "="*80)
    tests = [("Chrome Installation", check_chrome_installation), ("Headless Resolution", check_headless_resolution), ("Visible Chrome Launch", check_visible_chrome_launch), ("Selenium Attachment", check_selenium_attachment)]
    results = [(name, test_func()) for name, test_func in tests]
    print_section("TEST SUMMARY", "-")
    for name, result in results:
        print(f"{'[PASS]' if result else '[FAIL]'} {name}")
    print_section("RECOMMENDATIONS", "=")
    print("""To run SOMA with visible browser:
1. Set HEADLESS=false in your .env file
2. Run: python -m soma_app.workflows.run_soma
3. Chrome window should appear
4. Monitor progress in the browser

Known fixes applied:
- _launch_visible_chrome(): Launches Chrome separately with remote-debugging-port
- _bring_window_to_foreground_by_pid(): Ensures window visibility
- debuggerAddress: Connects Selenium to running Chrome instance
- Timeout: 45 seconds for remote port to become accessible

Troubleshooting:
- Verify Chrome.exe is installed at: C:\Program Files\Google\Chrome\Application\chrome.exe
- Check HEADLESS environment variable is 'false' or 'no'
- Use Task Manager (tasklist /v) to see Chrome windows
- If hidden, try setting these env vars: HEADLESS=false DISPLAY=:0""")
    all_passed = all(r for _, r in results)
    print_section("CONCLUSION", "=")
    status = "[SUCCESS] All tests passed!" if all_passed else "[FAILURE] Some tests failed"
    print(status)
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())
