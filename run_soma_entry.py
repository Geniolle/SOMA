from __future__ import annotations

import os
import sys
from pathlib import Path


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _base_dir() -> Path:
    # Em modo executável: pasta onde está o .exe
    if _is_frozen() and hasattr(sys, "executable"):
        return Path(sys.executable).resolve().parent
    # Em dev: pasta do projeto
    return Path(__file__).resolve().parent


def _first_existing(*paths: Path) -> Path | None:
    for p in paths:
        try:
            if p and p.exists():
                return p
        except OSError:
            pass
    return None


def _parse_env_file_for_key(env_file: Path, key: str) -> str | None:
    try:
        for line in env_file.read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if not s or s.startswith("#") or "=" not in s:
                continue
            k, v = s.split("=", 1)
            if k.strip() == key:
                return v.strip().strip('"').strip("'")
    except Exception:
        return None
    return None


def _set_default_env_files() -> None:
    base = _base_dir()
    internal = base / "_internal"

    # 1) ENV_FILE (config.env)
    if not (os.getenv("ENV_FILE") or "").strip():
        env_file = _first_existing(
            base / "deploy" / "config.env",
            internal / "deploy" / "config.env",
        )
        if env_file:
            os.environ["ENV_FILE"] = str(env_file)

    # 2) GOOGLE_CREDENTIALS_PATH
    if not (os.getenv("GOOGLE_CREDENTIALS_PATH") or "").strip():
        env_file_val = (os.getenv("ENV_FILE") or "").strip()
        env_file = Path(env_file_val) if env_file_val else None

        # tenta ler do config.env (ex: deploy\credenciais.json)
        rel_from_env = None
        if env_file and env_file.exists():
            rel_from_env = _parse_env_file_for_key(env_file, "GOOGLE_CREDENTIALS_PATH")

        candidates: list[Path] = []

        if rel_from_env:
            rel_path = Path(rel_from_env)
            # tenta relativo ao base e ao _internal
            candidates += [
                (base / rel_path),
                (internal / rel_path),
            ]

        # fallbacks comuns (caso config.env não tenha a chave ou esteja diferente)
        candidates += [
            base / "deploy" / "credenciais.json",
            internal / "deploy" / "credenciais.json",
            base / "secrets" / "credenciais.json",
            internal / "secrets" / "credenciais.json",
            base / "deploy" / "secrets" / "credenciais.json",
            internal / "deploy" / "secrets" / "credenciais.json",
            base / "credenciais.json",
            internal / "credenciais.json",
        ]

        cred_file = _first_existing(*candidates)
        if cred_file:
            os.environ["GOOGLE_CREDENTIALS_PATH"] = str(cred_file)

    # 3) PYTHONPATH só ajuda em dev (no exe não é necessário)
    if not _is_frozen() and not (os.getenv("PYTHONPATH") or "").strip():
        src = base / "src"
        if src.exists():
            os.environ["PYTHONPATH"] = str(src)


def main() -> int:
    _set_default_env_files()
    from soma_app.workflows.run_soma import main as run_main
    return int(run_main() or 0)


if __name__ == "__main__":
    raise SystemExit(main())