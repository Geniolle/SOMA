"""Diagnostic tool to reproduce and analyze headless Chrome issue."""
import os, sys, time, platform, subprocess
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

def check_chrome():
    from soma_app.infra.webdriver_factory import _find_chrome_executable
    try:
        chrome_exe = _find_chrome_executable()
        print(f"[OK] Chrome found: {chrome_exe}")
        result = subprocess.run([chrome_exe, '--version'], capture_output=True, text=True, timeout=5)
        print(f"[OK] Version: {result.stdout.strip() or result.stderr.strip()}")
        return True
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def check_headless():
    os.environ['HEADLESS'] = 'false'
    from soma_app.infra.webdriver_factory import _resolve_headless
    r1 = _resolve_headless(headless=False)
    r2 = _resolve_headless(headless=True)
    r3 = _resolve_headless()
    print(f"[OK] Headless logic: False={r1}, True={r2}, Env={r3}")
    return not r1 and r2 and not r3

def check_chrome_launch():
    from soma_app.infra.webdriver_factory import _launch_visible_chrome
    try:
        port, profile_dir, proc = _launch_visible_chrome()
        print(f"[OK] Chrome launched: PID={proc.pid}, Port={port}")
        if proc.poll() is None:
            time.sleep(1)
            proc.terminate()
            proc.wait(timeout=5)
            print("[OK] Chrome terminated")
            return True
        return False
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

def check_selenium():
    os.environ.update({'GOOGLE_CREDENTIALS_PATH': 'deploy/credenciais.json', 'SPREADSHEET_URL': 'https://test', 'SITE_USER': 'test', 'SITE_PASSWORD': 'test'})
    from soma_app.infra.webdriver_factory import create_driver
    try:
        driver = create_driver(headless=False)
        print(f"[OK] Driver created: Handles={len(driver.window_handles)}")
        driver.quit()
        return True
    except Exception as e:
        print(f"[ERROR] {e}")
        return False

print("=== SOMA Headless Chrome Diagnostic ===")
tests = [("Chrome", check_chrome), ("Headless", check_headless), ("Launch", check_chrome_launch), ("Selenium", check_selenium)]
results = [(name, test()) for name, test in tests]
print("\n=== Results ===")
for name, r in results:
    print(f"{'PASS' if r else 'FAIL'}: {name}")
print("\n=== Recommendations ===")
print("1. Set HEADLESS=false in .env file")
print("2. Run: python -m soma_app.workflows.run_soma")
print("3. Chrome window should appear")
print("4. Fixes: _launch_visible_chrome(), debuggerAddress, _bring_window_to_foreground_by_pid()")
