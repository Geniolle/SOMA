# scaffold_soma.py
# Cria automaticamente a estrutura de pastas (e ficheiros "placeholder") do projeto SOMA.
# Compatível com Windows/macOS/Linux.

from __future__ import annotations

import argparse
from pathlib import Path


DIRS = [
    "src/soma_app",
    "src/soma_app/config",
    "src/soma_app/infra",
    "src/soma_app/automation",
    "src/soma_app/automation/pages",
    "src/soma_app/domain",
    "src/soma_app/workflows",
    "src/soma_app/ui",
    "tests",
]

FILES = [
    "README.md",
    "requirements.txt",
    ".env.example",
    "src/soma_app/__init__.py",
    "src/soma_app/main.py",
    "src/soma_app/config/__init__.py",
    "src/soma_app/config/settings.py",
    "src/soma_app/infra/__init__.py",
    "src/soma_app/infra/logging.py",
    "src/soma_app/infra/sheets_client.py",
    "src/soma_app/infra/webdriver_factory.py",
    "src/soma_app/automation/__init__.py",
    "src/soma_app/automation/actions.py",
    "src/soma_app/automation/selectors.py",
    "src/soma_app/automation/pages/__init__.py",
    "src/soma_app/automation/pages/login_page.py",
    "src/soma_app/automation/pages/entradas_saidas_page.py",
    "src/soma_app/automation/pages/transferencias_page.py",
    "src/soma_app/automation/pages/caixas_bancos_page.py",
    "src/soma_app/domain/__init__.py",
    "src/soma_app/domain/models.py",
    "src/soma_app/domain/rules.py",
    "src/soma_app/workflows/__init__.py",
    "src/soma_app/workflows/process_contaordem.py",
    "src/soma_app/workflows/update_caixas.py",
    "src/soma_app/ui/__init__.py",
    "src/soma_app/ui/notifier.py",
    "tests/test_rules.py",
    "tests/test_parsing.py",
]

TEMPLATES = {
    "README.md": "# SOMA\n\nEstrutura inicial do projeto.\n",
    "requirements.txt": "# Adiciona aqui as dependências (ex.: selenium, gspread, oauth2client, python-dotenv)\n",
    ".env.example": (
        "GOOGLE_CREDENTIALS_PATH=\n"
        "SPREADSHEET_URL=\n"
        "SHEET_CONTAORDEM=CONTAORDEM\n"
        "SHEET_CAIXAS=GERENCIAR CAIXAS\n"
        "SITE_USER=\n"
        "SITE_PASSWORD=\n"
        "HEADLESS=true\n"
        "TIMEOUT_SECONDS=20\n"
        "RETRY_COUNT=2\n"
        "LOG_LEVEL=INFO\n"
        "RUN_ENV=dev\n"
    ),
    "src/soma_app/main.py": (
        '"""Entry point do SOMA.\n\n'
        "Mantém este ficheiro pequeno: carregar settings, configurar logging e chamar workflows.\n"
        '"""\n\n'
        "def main() -> int:\n"
        "    return 0\n\n"
        "if __name__ == '__main__':\n"
        "    raise SystemExit(main())\n"
    ),
}


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def ensure_file(path: Path, force: bool = False) -> None:
    if path.exists() and not force:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    content = TEMPLATES.get(str(path.as_posix()), "")
    path.write_text(content, encoding="utf-8")


def scaffold(root: Path, create_files: bool, force_files: bool) -> None:
    created_dirs = 0
    created_files = 0

    for d in DIRS:
        p = root / d
        if not p.exists():
            created_dirs += 1
        ensure_dir(p)

    if create_files:
        for f in FILES:
            p = root / f
            before = p.exists()
            ensure_file(p, force=force_files)
            after = p.exists()
            if not before and after:
                created_files += 1
            elif force_files and after:
                created_files += 1

    print(f"[OK] Root: {root.resolve()}")
    print(f"[OK] Pastas garantidas: {len(DIRS)} (novas: {created_dirs})")
    if create_files:
        print(f"[OK] Ficheiros garantidos: {len(FILES)} (criados/sobrescritos: {created_files})")
    else:
        print("[OK] Ficheiros: não criados (usa --files para criar placeholders)")


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Scaffold da estrutura SOMA")
    ap.add_argument("--root", default=".", help="Pasta raiz do projeto (default: .)")
    ap.add_argument("--files", action="store_true", help="Cria também ficheiros placeholder")
    ap.add_argument("--force-files", action="store_true", help="Sobrescreve ficheiros existentes (cuidado)")
    return ap.parse_args()


if __name__ == "__main__":
    args = parse_args()
    scaffold(Path(args.root), create_files=args.files, force_files=args.force_files)
