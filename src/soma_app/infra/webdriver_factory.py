# src/soma_app/infra/webdriver_factory.py
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import ctypes
import socket
import tempfile
import time
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterator, Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from soma_app.infra.env import env_bool
from soma_app.infra.log_config import ensure_artifacts_dirs
from soma_app.infra.trace import log_kv

logger = logging.getLogger(__name__)


def _get_setting(settings: Any, *names: str, default: Any = None) -> Any:
    for n in names:
        if settings is not None and hasattr(settings, n):
            v = getattr(settings, n)
            if v is not None and v != "":
                return v
    return default


def _resolve_headless(settings: Any = None, headless: Optional[bool] = None) -> bool:
    if headless is not None:
        return bool(headless)

    v = _get_setting(settings, "headless", "HEADLESS", default=None)
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        return v.strip().lower() in {"1", "true", "yes", "y", "on"}

    return env_bool("HEADLESS", default=True)


def _resolve_downloads_dir(settings: Any = None, downloads_dir: Optional[str] = None) -> str:
    if downloads_dir:
        return os.fspath(downloads_dir)

    paths = ensure_artifacts_dirs(settings)
    return paths["downloads_dir"]


def _find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _find_chrome_executable() -> str:
    candidates = [
        shutil.which("chrome"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    raise FileNotFoundError("chrome.exe não encontrado.")


def _launch_visible_chrome() -> tuple[int, str, subprocess.Popen[str]]:
    chrome_exe = _find_chrome_executable()
    port = _find_free_port()

    temp_root = Path(os.environ.get("TMP", r"C:\tmp"))
    temp_root.mkdir(parents=True, exist_ok=True)
    profile_dir = tempfile.mkdtemp(prefix="soma_chrome_", dir=str(temp_root))

    args = [
        chrome_exe,
        "--new-window",
        f"--remote-debugging-port={port}",
        "--remote-allow-origins=*",
        f"--user-data-dir={profile_dir}",
        "--start-maximized",
        "--window-position=0,0",
        "--window-size=1920,1080",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-background-networking",
        "--disable-sync",
        "--disable-notifications",
        "--disable-default-apps",
        "--disable-features=PushMessaging,MediaRouter,Translate",
        "about:blank",
    ]

    proc = subprocess.Popen(
        args,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
    )

    deadline = time.time() + 45
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1.0) as resp:
                body = resp.read().decode("utf-8", "ignore")
                if "webSocketDebuggerUrl" in body:
                    time.sleep(5)
                    break
        except Exception:
            if proc.poll() is not None:
                raise RuntimeError("Chrome terminou antes de abrir a porta de debugging.")
            time.sleep(0.2)

    return port, profile_dir, proc


def _build_options(headless: bool, downloads_dir: str) -> Options:
    opt = Options()

    # Headless/viewport
    if headless:
        opt.add_argument("--headless=new")
    # === BLINDAGEM CONTRA VERSÃO MOBILE / HEADLESS ===
    opt.add_argument("--start-maximized")
    opt.add_argument("--window-position=0,0")
    opt.add_argument("--window-size=1920,1080")
    opt.add_argument("--force-device-scale-factor=1")
    opt.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    opt.add_argument("--disable-gpu")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-dev-shm-usage")

    # “Calar” logs do Chrome (reduz bastante no console)
    opt.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opt.add_experimental_option("useAutomationExtension", False)
    opt.add_argument("--log-level=3")

    # Reduz serviços de background (ajuda a diminuir GCM/PUSH noise)
    opt.add_argument("--disable-background-networking")
    opt.add_argument("--disable-sync")
    opt.add_argument("--disable-notifications")
    opt.add_argument("--disable-default-apps")
    opt.add_argument("--no-first-run")
    opt.add_argument("--no-default-browser-check")
    opt.add_argument("--disable-component-update")
    opt.add_argument("--disable-breakpad")
    opt.add_argument("--disable-crash-reporter")
    opt.add_argument("--metrics-recording-only")
    opt.add_argument("--disable-client-side-phishing-detection")

    # Algumas features que costumam gerar chatter
    opt.add_argument("--disable-features=PushMessaging,MediaRouter,Translate")

    # Downloads
    prefs = {
        "download.default_directory": os.path.abspath(downloads_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opt.add_experimental_option("prefs", prefs)

    return opt


def _build_service() -> Service:
    # Silencia logs do ChromeDriver (não é o mesmo que logs do Chrome, mas ajuda)
    try:
        return Service(log_output=os.devnull)
    except TypeError:
        return Service()


def _bring_window_to_foreground_by_pid(pid: int) -> bool:
    if not pid:
        return False

    user32 = ctypes.windll.user32
    handles: list[int] = []

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

    def _enum_cb(hwnd, lparam):
        proc_id = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
        if proc_id.value != pid:
            return True
        if user32.IsWindowVisible(hwnd):
            handles.append(int(hwnd))
        return True

    try:
        user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
    except Exception:
        return False

    if not handles:
        return False

    hwnd = handles[0]
    try:
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    except Exception:
        pass
    try:
        user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040)
    except Exception:
        pass
    try:
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass
    try:
        user32.BringWindowToTop(hwnd)
    except Exception:
        pass
    return True


@dataclass
class WebDriverBundle:
    """
    Bundle compatível:
      - pode ser usado como driver (proxy via __getattr__)
      - expõe driver/a/downloads_dir
      - permite unpack (driver, a, downloads_dir)
    """
    driver: webdriver.Chrome
    a: Any
    downloads_dir: str

    def __getattr__(self, item: str) -> Any:
        return getattr(self.driver, item)

    def __iter__(self) -> Iterator[Any]:
        yield self.driver
        yield self.a
        yield self.downloads_dir

    def quit(self) -> None:
        try:
            self.driver.quit()
        except Exception:
            pass


def _create_actions(driver: webdriver.Chrome) -> Any:
    """
    Cria o wrapper de actions do projeto (se existir).
    """
    try:
        from soma_app.automation.actions import ActionConfig, Actions  # type: ignore

        # Actions agora exige cfg: ActionConfig
        return Actions(driver, ActionConfig())
    except Exception as e:
        # Mantém comportamento "best effort" (não quebra create_bundle),
        # mas deixa rasto para diagnóstico.
        logger.exception("Falha ao criar Actions(driver, ActionConfig()): %s", e)
        return None


def create_driver(
    settings: Any = None,
    *,
    headless: Optional[bool] = None,
    downloads_dir: Optional[str] = None,
) -> webdriver.Chrome:
    headless_v = _resolve_headless(settings, headless)
    downloads_v = _resolve_downloads_dir(settings, downloads_dir)

    service = _build_service()
    if headless_v:
        options = _build_options(headless=headless_v, downloads_dir=downloads_v)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        port, profile_dir, chrome_proc = _launch_visible_chrome()
        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
        driver = webdriver.Chrome(service=service, options=options)
        log_kv(
            logger,
            "Chrome visible launcher",
            pid=chrome_proc.pid,
            port=port,
            profile_dir=profile_dir,
        )
    
    # === FORÇAR MAXIMIZAÇÃO (Dupla Segurança) ===
    try:
        driver.set_window_rect(0, 0, 1920, 1080)
    except Exception:
        pass
    try:
        driver.maximize_window()
    except Exception:
        pass

    try:
        caps = getattr(driver, "capabilities", {}) or {}
        browser_pid = int(caps.get("goog:processID") or caps.get("processID") or 0)
    except Exception:
        browser_pid = 0

    if browser_pid:
        try:
            found_window = _bring_window_to_foreground_by_pid(browser_pid)
            log_kv(logger, "Chrome foreground probe", browser_pid=browser_pid, found_window=found_window)
        except Exception:
            pass

    # headless downloads via CDP (best effort)
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": os.path.abspath(downloads_v)},
        )
    except Exception:
        pass

    log_kv(logger, "WebDriver OK", headless=headless_v, downloads=downloads_v)
    return driver


def create_bundle(
    settings: Any = None,
    *,
    headless: Optional[bool] = None,
    downloads_dir: Optional[str] = None,
) -> WebDriverBundle:
    downloads_v = _resolve_downloads_dir(settings, downloads_dir)
    driver = create_driver(settings, headless=headless, downloads_dir=downloads_v)
    actions = _create_actions(driver)
    return WebDriverBundle(driver=driver, a=actions, downloads_dir=downloads_v)


# Aliases para compatibilidade
def create_webdriver(settings: Any = None, *, headless: Optional[bool] = None, downloads_dir: Optional[str] = None):
    return create_bundle(settings, headless=headless, downloads_dir=downloads_dir)


def build_driver(settings: Any = None, *, headless: Optional[bool] = None, downloads_dir: Optional[str] = None):
    return create_bundle(settings, headless=headless, downloads_dir=downloads_dir)


def get_driver(settings: Any = None, *, headless: Optional[bool] = None, downloads_dir: Optional[str] = None):
    return create_bundle(settings, headless=headless, downloads_dir=downloads_dir)


class WebDriverFactory:
    """
    Compatível com o teu uso atual:
      bundle = WebDriverFactory.create(settings)
    """

    @staticmethod
    def create(
        settings: Any = None,
        *,
        headless: Optional[bool] = None,
        downloads_dir: Optional[str] = None,
    ) -> WebDriverBundle:
        return create_bundle(settings, headless=headless, downloads_dir=downloads_dir)

    def __init__(self, settings: Any = None):
        self._settings = settings

    def create_instance(self, *, headless: Optional[bool] = None, downloads_dir: Optional[str] = None) -> WebDriverBundle:
        return create_bundle(self._settings, headless=headless, downloads_dir=downloads_dir)


# -----------------------------------------------------------------------------
# ChromeDriver version diagnostics
# -----------------------------------------------------------------------------
def unwrap_webdriver(obj: Any) -> Any:
    """
    bundle.a pode ser um wrapper. Tenta chegar ao webdriver real.
    """
    cur = obj
    seen = set()
    for _ in range(6):
        if cur is None:
            return None
        oid = id(cur)
        if oid in seen:
            return cur
        seen.add(oid)

        if hasattr(cur, "capabilities") and (hasattr(cur, "execute_script") or hasattr(cur, "execute_cdp_cmd")):
            return cur

        for attr in ("driver", "_driver", "webdriver", "_webdriver", "wd", "_wd", "browser", "_browser"):
            nxt = getattr(cur, attr, None)
            if nxt is not None and nxt is not cur:
                cur = nxt
                break
        else:
            return cur
    return cur


def _get_driver_path(driver: Any) -> str:
    d = unwrap_webdriver(driver)
    try:
        svc = getattr(d, "service", None) or getattr(d, "_service", None)
        p = getattr(svc, "path", None)
        if isinstance(p, str) and p.strip():
            return p.strip()
    except Exception:
        pass
    return ""


def _chromedriver_version_from_exe(driver_path: str) -> str:
    p = (driver_path or "").strip()
    if not p:
        return ""
    try:
        r = subprocess.run([p, "--version"], capture_output=True, text=True, timeout=5)
        out = (r.stdout or r.stderr or "").strip()
        m = re.search(r"ChromeDriver\s+(\d+\.\d+\.\d+\.\d+)", out)
        return m.group(1) if m else out[:120]
    except Exception:
        return ""


def _chromedriver_version_from_capabilities(driver: Any) -> str:
    d = unwrap_webdriver(driver)
    try:
        caps = getattr(d, "capabilities", {}) or {}
        if isinstance(caps, dict):
            chrome = caps.get("chrome")
            if isinstance(chrome, dict):
                v = chrome.get("chromedriverVersion")
                if isinstance(v, str) and v.strip():
                    return v.split(" ")[0].strip()
    except Exception:
        pass
    return ""


def _find_chromedriver_in_known_caches() -> str:
    """
    Fallback: procura chromedriver.exe nos caches comuns:
      - Selenium Manager: %USERPROFILE%\\.cache\\selenium\\...
      - webdriver_manager: %USERPROFILE%\\.wdm\\...
    """
    candidates: list[Path] = []

    home = Path.home()
    roots = [
        home / ".cache" / "selenium",
        home / ".wdm",
    ]

    la = os.getenv("LOCALAPPDATA")
    if la:
        roots.append(Path(la) / "selenium")
    tmp = os.getenv("TEMP")
    if tmp:
        roots.append(Path(tmp) / "selenium")

    for root in roots:
        try:
            if not root.exists():
                continue
            for p in root.rglob("chromedriver.exe"):
                candidates.append(p)
        except Exception:
            continue

    if not candidates:
        return ""

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])


def get_chrome_version(driver: Any) -> str:
    d = unwrap_webdriver(driver)

    try:
        caps = getattr(d, "capabilities", {}) or {}
        if isinstance(caps, dict):
            v = caps.get("browserVersion")
            if isinstance(v, str) and v.strip():
                return v.strip()
    except Exception:
        pass

    try:
        if hasattr(d, "execute_cdp_cmd"):
            info = d.execute_cdp_cmd("Browser.getVersion", {})
            if isinstance(info, dict):
                prod = info.get("product")
                if isinstance(prod, str) and prod.strip():
                    return prod.strip()
    except Exception:
        pass

    return ""


def get_chromedriver_info(driver: Any) -> Dict[str, str]:
    d = unwrap_webdriver(driver)

    v = _chromedriver_version_from_capabilities(d)
    if v:
        return {"version": v, "path": _get_driver_path(d) or "n/a", "source": "capabilities"}

    p = _get_driver_path(d)
    if p:
        v2 = _chromedriver_version_from_exe(p)
        if v2:
            return {"version": v2, "path": p, "source": "service.path"}

    p3 = _find_chromedriver_in_known_caches()
    if p3:
        v3 = _chromedriver_version_from_exe(p3)
        if v3:
            return {"version": v3, "path": p3, "source": "cache"}

    return {"version": "", "path": p or "n/a", "source": "unknown"}
