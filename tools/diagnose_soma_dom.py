from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any, Iterable

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.config.settings import Settings
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.webdriver_factory import WebDriverFactory

log = logging.getLogger("soma_app.tools.diagnose_soma_dom")


def _min_row() -> ContaOrdemRow:
    return ContaOrdemRow(
        row_number=999999,
        tipo=TipoMovimento.ENTRADA,
        data_mov="21/07/2026",
        caixa="",
        caixa_saida="",
        centro_custo="",
        plano_conta="",
        forma_pagamento="",
        importancia="",
        descricao_soma="",
        doc_soma="",
        dados_doc="",
        iduser="",
        timestamp="",
        raw={},
    )


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


def _safe_name(value: Any) -> str:
    s = "" if value is None else str(value)
    s = " ".join(s.split())
    s = s.strip()
    return s


def _candidate_xpath_from_attrs(tag: str, attrs: dict[str, str]) -> list[str]:
    out: list[str] = []
    tag = tag or "*"
    for attr in ("id", "name", "placeholder", "aria-label"):
        val = _safe_name(attrs.get(attr))
        if val:
            out.append(f"//{tag}[@{attr}={json.dumps(val)}]")
    text = _safe_name(attrs.get("text"))
    if text and tag in {"button", "a", "span", "label"}:
        out.append(f"//{tag}[contains(normalize-space(.), {json.dumps(text[:80])})]")
    return out


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
        placeholder: el.getAttribute('placeholder') || '',
        aria_label: el.getAttribute('aria-label') || '',
        role: el.getAttribute('role') || '',
        text: (el.innerText || el.textContent || '').trim(),
        value: el.value || '',
        displayed: !!(rect.width || rect.height) && style.display !== 'none' && style.visibility !== 'hidden',
        x: Math.round(rect.x),
        y: Math.round(rect.y),
        w: Math.round(rect.width),
        h: Math.round(rect.height),
        outer_html: el.outerHTML,
      };
    });
    """
    return driver.execute_script(script)


def _find_desc_candidates(elements: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    matches = []
    needles = ("descricao", "descrição", "descriçao", "description")
    for el in elements:
        blob = " ".join(
            _safe_name(el.get(k))
            for k in ("tag", "type", "name", "id", "class", "placeholder", "aria_label", "role", "text", "value")
        ).lower()
        if any(n in blob for n in needles):
            attrs = {
                "id": _safe_name(el.get("id")),
                "name": _safe_name(el.get("name")),
                "placeholder": _safe_name(el.get("placeholder")),
                "aria-label": _safe_name(el.get("aria_label")),
                "text": _safe_name(el.get("text")),
            }
            matches.append(
                {
                    "element": el,
                    "xpath_candidates": _candidate_xpath_from_attrs(str(el.get("tag") or "*"), attrs),
                }
            )
    return matches


def main() -> int:
    settings = _load_settings()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    bundle = WebDriverFactory.create(settings, headless=False)
    driver = bundle.driver
    try:
        login = LoginPage(bundle.a, settings)
        login.login()

        page = EntradasSaidasPage(bundle.a, settings)
        probe_row = _min_row()
        page._open_new(probe_row)
        page._choose_tipo(probe_row)

        out_dir = Path(settings.screenshots_dir).parent / "diagnostics"
        out_dir.mkdir(parents=True, exist_ok=True)

        source_path = bundle.a.dump_page_source("soma_nova_entrada_saida")
        screenshot_path = bundle.a.screenshot("soma_nova_entrada_saida")
        interactive = _scan_interactive(driver)
        desc_hits = _find_desc_candidates(interactive)

        interactive_path = out_dir / "soma_nova_entrada_saida_interactive.json"
        desc_path = out_dir / "soma_nova_entrada_saida_descricao_candidates.json"
        summary_path = out_dir / "soma_nova_entrada_saida_summary.json"

        interactive_path.write_text(json.dumps(interactive, ensure_ascii=False, indent=2), encoding="utf-8")
        desc_path.write_text(json.dumps(desc_hits, ensure_ascii=False, indent=2), encoding="utf-8")
        summary_path.write_text(
            json.dumps(
                {
                    "url": driver.current_url,
                    "title": driver.title,
                    "page_source": str(source_path),
                    "screenshot": str(screenshot_path),
                    "interactive_elements": len(interactive),
                    "descricao_hits": len(desc_hits),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        log.warning("Diagnóstico concluído | url=%s | title=%s | elements=%s | desc_hits=%s", driver.current_url, driver.title, len(interactive), len(desc_hits))
        log.warning("Artefactos: %s | %s | %s", interactive_path, desc_path, summary_path)
        return 0
    finally:
        try:
            bundle.quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
