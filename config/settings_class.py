from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv


def _to_bool(v: str, default: bool = False) -> bool:
    if v is None:
        return default
    s = str(v).strip().lower()
    return s in ("1", "true", "t", "yes", "y", "sim", "on")


def _to_int(v: str, default: int) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return default


def _strip_quotes(v: str) -> str:
    s = (v or "").strip()
    if len(s) >= 2 and ((s[0] == s[-1]) and s[0] in ("'", '"')):
        return s[1:-1].strip()
    return s


@dataclass(frozen=True)
class Settings:
    # Google / Sheets
    google_credentials_path: Path
    spreadsheet_url: str
    sheet_contaordem: str
    sheet_caixas: str

    # Site (Selenium)
    site_user: str
    site_password: str
    site_login_url: str
    site_home_url: str

    # Execução
    headless: bool
    timeout_seconds: int
    retry_count: int
    batch_size: int

    # Observabilidade
    log_level: str
    run_env: str
    log_dir: Path

    # Artefactos
    screenshots_dir: Path

    @classmethod
    def from_env(cls, env_file: Optional[str] = None) -> "Settings":
        load_dotenv(dotenv_path=env_file, override=False)

        cred_path = Path(_strip_quotes(os.getenv("GOOGLE_CREDENTIALS_PATH", "")))
        spreadsheet_url = _strip_quotes(os.getenv("SPREADSHEET_URL", ""))

        sheet_contaordem = _strip_quotes(os.getenv("SHEET_CONTAORDEM", "CONTAORDEM"))
        sheet_caixas = _strip_quotes(os.getenv("SHEET_CAIXAS", "GERENCIAR CAIXAS"))

        site_user = _strip_quotes(os.getenv("SITE_USER", ""))
        site_password = _strip_quotes(os.getenv("SITE_PASSWORD", ""))

        site_login_url = _strip_quotes(os.getenv("SITE_LOGIN_URL", "https://verbodavida.info/apps/index.php"))
        site_home_url = _strip_quotes(os.getenv("SITE_HOME_URL", "https://verbodavida.info/IVV/"))

        headless = _to_bool(os.getenv("HEADLESS", "true"), default=True)
        timeout_seconds = _to_int(os.getenv("TIMEOUT_SECONDS", "20"), default=20)
        retry_count = _to_int(os.getenv("RETRY_COUNT", "2"), default=2)
        batch_size = _to_int(os.getenv("BATCH_SIZE", "20"), default=20)

        log_level = _strip_quotes(os.getenv("LOG_LEVEL", "INFO")).upper()
        run_env = _strip_quotes(os.getenv("RUN_ENV", "dev"))

        log_dir = Path(_strip_quotes(os.getenv("LOG_DIR", "logs")))
        screenshots_dir = Path(_strip_quotes(os.getenv("SCREENSHOTS_DIR", "artifacts/screenshots")))

        settings = cls(
            google_credentials_path=cred_path,
            spreadsheet_url=spreadsheet_url,
            sheet_contaordem=sheet_contaordem,
            sheet_caixas=sheet_caixas,
            site_user=site_user,
            site_password=site_password,
            site_login_url=site_login_url,
            site_home_url=site_home_url,
            headless=headless,
            timeout_seconds=timeout_seconds,
            retry_count=retry_count,
            batch_size=batch_size,
            log_level=log_level,
            run_env=run_env,
            log_dir=log_dir,
            screenshots_dir=screenshots_dir,
        )
        settings._validate()
        return settings

    def _validate(self) -> None:
        missing = []

        if not str(self.google_credentials_path).strip():
            missing.append("GOOGLE_CREDENTIALS_PATH")
        if not self.spreadsheet_url.strip():
            missing.append("SPREADSHEET_URL")
        if not self.site_user.strip():
            missing.append("SITE_USER")
        if not self.site_password.strip():
            missing.append("SITE_PASSWORD")

        if missing:
            raise ValueError(f"Config em falta no .env/variáveis: {', '.join(missing)}")

        if not self.google_credentials_path.exists():
            raise FileNotFoundError(f"GOOGLE_CREDENTIALS_PATH não existe: {self.google_credentials_path}")
        if not self.google_credentials_path.is_file():
            raise ValueError(f"GOOGLE_CREDENTIALS_PATH deve ser ficheiro JSON: {self.google_credentials_path}")
        if self.google_credentials_path.suffix.lower() != ".json":
            raise ValueError(f"GOOGLE_CREDENTIALS_PATH deve ser .json: {self.google_credentials_path}")

        if self.timeout_seconds <= 0:
            raise ValueError("TIMEOUT_SECONDS deve ser > 0")
        if self.retry_count < 0:
            raise ValueError("RETRY_COUNT deve ser >= 0")
        if self.batch_size <= 0:
            raise ValueError("BATCH_SIZE deve ser > 0")