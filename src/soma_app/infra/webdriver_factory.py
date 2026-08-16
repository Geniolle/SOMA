# src/soma_app/infra/webdriver_factory.py
from __future__ import annotations

import logging
import os
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterator, Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.remote_connection import RemoteConnection

from soma_app.infra.env import env_bool, env_int
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


def _build_options(headless: bool, downloads_dir: str) -> Options:
    opt = Options()

    # Headless/viewport
    if headless:
        opt.add_argument("--headless=new")
    
    # === BLINDAGEM CONTRA VERSÃO MOBILE / HEADLESS ===
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
        return Actions(
            driver,
            ActionConfig(
                selector_debug_interactive=env_bool("DEBUG_SELECTOR_INTERACTIVE", default=False),
            ),
        )
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
    if env_bool("DEBUG_STEP_MODE", default=False) and headless_v:
        raise RuntimeError("DEBUG_STEP_MODE=true requer HEADLESS=false para manter o browser visível.")
    downloads_v = _resolve_downloads_dir(settings, downloads_dir)
    http_timeout = env_int("SELENIUM_HTTP_TIMEOUT", 300)

    options = _build_options(headless=headless_v, downloads_dir=downloads_v)
    service = _build_service()

    log_kv(logger, "WebDriver create start", headless=headless_v, downloads=downloads_v)
    previous_timeout = None
    try:
        try:
            previous_timeout = RemoteConnection.get_timeout()
        except Exception:
            previous_timeout = None
        RemoteConnection.set_timeout(http_timeout)
        driver = webdriver.Chrome(service=service, options=options)
    finally:
        if previous_timeout is not None:
            try:
                RemoteConnection.set_timeout(previous_timeout)
            except Exception:
                pass
    log_kv(logger, "WebDriver chrome ready", headless=headless_v, downloads=downloads_v)
    
    # === FORÇAR MAXIMIZAÇÃO (Dupla Segurança) ===
    if not headless_v:
        try:
            logger.info("WebDriver maximize best-effort")
            driver.maximize_window()
        except Exception:
            try:
                logger.info("WebDriver maximize falhou; aplicando set_window_size")
                driver.set_window_size(1920, 1080)
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
