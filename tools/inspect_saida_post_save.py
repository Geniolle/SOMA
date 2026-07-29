from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.config.settings import Settings
from soma_app.domain.models import ContaOrdemRow
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.sheets_client import SheetsClient
from soma_app.infra.webdriver_factory import WebDriverFactory

log = logging.getLogger("soma_app.tools.inspect_saida_post_save")


def _load_settings() -> Settings:
    project_root = Path(__file__).resolve().parents[1]
    env_file = (os.getenv("ENV_FILE") or "").strip()
    if env_file:
        env_path = Path(env_file)
        if not env_path.is_absolute():
            cwd_candidate = Path.cwd() / env_path
            if cwd_candidate.exists():
                return Settings.from_env(env_file=str(cwd_candidate))
            project_candidate = project_root / env_path
            if project_candidate.exists():
                return Settings.from_env(env_file=str(project_candidate))
        elif env_path.exists():
            return Settings.from_env(env_file=str(env_path))

    default_env = project_root / "deploy" / ".env"
    if default_env.exists():
        return Settings.from_env(env_file=str(default_env))
    return Settings.from_env(env_file=None)


def _scan_interactive(driver) -> list[dict[str, Any]]:
    script = """
    const nodes = Array.from(document.querySelectorAll('input, textarea, select, button, a, [role="button"], [contenteditable="true"]'));
    return nodes.map((el, index) => {
      const rect = el.getBoundingClientRect();
      const style = window.getComputedStyle(el);
      return {
        index,
        tag: el.tagName.toLowerCase(),
        type: el.getAttribute('type') || '',
        name: el.getAttribute('name') || '',
        id: el.getAttribute('id') || '',
        class: el.getAttribute('class') || '',
        text: (el.innerText || el.textContent || '').trim(),
        value: el.value || '',
        displayed: !!(rect.width || rect.height) && style.display !== 'none' && style.visibility !== 'hidden',
        outer_html: el.outerHTML,
      };
    });
    """
    return driver.execute_script(script)


def main() -> int:
    settings = _load_settings()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    sheets = SheetsClient(settings)
    row_raw = sheets.get_all_records("CONTAORDEM")[1]
    row = ContaOrdemRow.from_table_row(row_number=2, raw=row_raw)

    bundle = WebDriverFactory.create(settings, headless=False)
    driver = bundle.driver
    try:
        login = LoginPage(bundle.a, settings)
        login.login()

        page = EntradasSaidasPage(bundle.a, settings)
        page._open_new(row)
        page._choose_tipo(row)
        page._fill_common(row)
        page._fill_saida(row)
        page._save_form_if_present(row)

        out_dir = Path(settings.screenshots_dir).parent / "diagnostics"
        out_dir.mkdir(parents=True, exist_ok=True)
        source_path = bundle.a.dump_page_source("saida_post_save_row_2")
        screenshot_path = bundle.a.screenshot("saida_post_save_row_2")
        interactive = _scan_interactive(driver)
        path = out_dir / "saida_post_save_row_2_interactive.json"
        path.write_text(json.dumps(interactive, ensure_ascii=False, indent=2), encoding="utf-8")

        log.warning(
            "Inspect concluído | url=%s | title=%s | source=%s | screenshot=%s | elements=%s",
            driver.current_url,
            driver.title,
            source_path,
            screenshot_path,
            len(interactive),
        )
        return 0
    finally:
        try:
            bundle.quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
