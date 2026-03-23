# src/soma_app/infra/webdriver_factory.py
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Any, Iterator, Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from soma_app.infra.log_config import ensure_artifacts_dirs
from soma_app.infra.trace import log_kv

logger = logging.getLogger(__name__)


def _bool_env(name: str, default: bool = False) -> bool:
    v = (os.getenv(name) or "").strip().lower()
    if v in {"1", "true", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "no", "n", "off"}:
        return False
    return default


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

    return _bool_env("HEADLESS", default=True)


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
        from soma_app.automation.actions import Actions, ActionConfig  # type: ignore

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

    options = _build_options(headless=headless_v, downloads_dir=downloads_v)
    service = _build_service()

    driver = webdriver.Chrome(service=service, options=options)
    
    # === FORÇAR MAXIMIZAÇÃO (Dupla Segurança) ===
    driver.maximize_window()

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