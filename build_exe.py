from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable

APP_NAME = "SOMA"
ENTRYPOINT = Path("src/soma_app/workflows/run_soma.py")
SRC_DIR = Path("src")
DIST_DIR = Path("dist")
BUILD_DIR = Path("build")
SPEC_FILE = Path(f"{APP_NAME}.spec")
RUNTIME_DIRS = ["artifacts", "logs"]
COPY_FILES = [
    ".env",
    "config.env",
    "credenciais.json",
    "run_soma.bat",
]
COPY_DIRS = [
    "secrets",
]


def _print_header(title: str) -> None:
    print()
    print("=" * 78)
    print(title)
    print("=" * 78)


def _run(cmd: list[str], cwd: Path) -> None:
    print(f"> {' '.join(cmd)}")
    subprocess.run(cmd, cwd=str(cwd), check=True)


def _remove_path(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path, ignore_errors=True)
    else:
        path.unlink(missing_ok=True)


def _safe_copy_file(src: Path, dst_dir: Path) -> None:
    if not src.exists() or not src.is_file():
        print(f"[IGNORADO] Ficheiro não encontrado: {src}")
        return
    dst = dst_dir / src.name
    shutil.copy2(src, dst)
    print(f"[OK] Ficheiro copiado: {src} -> {dst}")


def _safe_copy_dir(src: Path, dst_dir: Path) -> None:
    if not src.exists() or not src.is_dir():
        print(f"[IGNORADO] Pasta não encontrada: {src}")
        return
    dst = dst_dir / src.name
    if dst.exists():
        shutil.rmtree(dst, ignore_errors=True)
    shutil.copytree(src, dst)
    print(f"[OK] Pasta copiada: {src} -> {dst}")


def _ensure_runtime_dirs(dist_app_dir: Path, names: Iterable[str]) -> None:
    for name in names:
        path = dist_app_dir / name
        path.mkdir(parents=True, exist_ok=True)
        print(f"[OK] Pasta runtime garantida: {path}")


def _write_launcher_bat(dist_app_dir: Path) -> None:
    bat_path = dist_app_dir / "iniciar_soma.bat"
    exe_name = f"{APP_NAME}.exe"
    content = f"""@echo off
setlocal
cd /d "%~dp0"
if not exist "{exe_name}" (
  echo ERRO: nao encontrei o executavel {exe_name} nesta pasta.
  pause
  exit /b 1
)
"{exe_name}"
set EXIT_CODE=%ERRORLEVEL%
echo.
echo Processo terminado com codigo: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
"""
    bat_path.write_text(content, encoding="utf-8", newline="\r\n")
    print(f"[OK] Launcher criado: {bat_path}")


def _write_readme(dist_app_dir: Path) -> None:
    readme = dist_app_dir / "LEIA-ME.txt"
    content = """
SOMA - Pacote do executável

Conteúdo esperado nesta pasta:
- SOMA.exe
- iniciar_soma.bat
- .env / config.env / credenciais.json (quando existirem no projeto)
- pasta secrets (quando existir no projeto)
- pasta artifacts
- pasta logs

Como executar:
1. Abra a pasta.
2. Dê duplo clique em iniciar_soma.bat.
3. O .bat executa o SOMA.exe e mantém a janela aberta no final.

Sempre que houver alterações no código:
1. Volte à raiz do projeto.
2. Execute: python build_exe.py
3. Recolha a nova pasta dist.
""".strip()
    readme.write_text(content, encoding="utf-8")
    print(f"[OK] LEIA-ME criado: {readme}")


def _find_built_exe(dist_dir: Path) -> Path:
    exe = dist_dir / f"{APP_NAME}.exe"
    if not exe.exists():
        raise FileNotFoundError(f"Executável não encontrado após build: {exe}")
    return exe


def main() -> int:
    root = Path(__file__).resolve().parent
    entrypoint = root / ENTRYPOINT
    src_dir = root / SRC_DIR
    dist_dir = root / DIST_DIR
    build_dir = root / BUILD_DIR
    spec_file = root / SPEC_FILE

    _print_header("VALIDAÇÃO INICIAL")
    print(f"Projeto raiz : {root}")
    print(f"Entrypoint   : {entrypoint}")
    print(f"Src dir      : {src_dir}")
    print(f"Dist dir     : {dist_dir}")
    print(f"Build dir    : {build_dir}")

    if not entrypoint.exists():
        raise FileNotFoundError(f"Entrypoint não encontrado: {entrypoint}")
    if not src_dir.exists():
        raise FileNotFoundError(f"Pasta src não encontrada: {src_dir}")

    _print_header("GARANTINDO PYINSTALLER")
    _run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], cwd=root)
    _run([sys.executable, "-m", "pip", "install", "pyinstaller"], cwd=root)

    _print_header("LIMPEZA DE BUILD ANTERIOR")
    _remove_path(dist_dir)
    _remove_path(build_dir)
    _remove_path(spec_file)
    print("[OK] Limpeza concluída.")

    _print_header("CRIANDO EXECUTÁVEL")
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name",
        APP_NAME,
        "--paths",
        str(src_dir),
        "--collect-submodules",
        "soma_app",
        str(entrypoint),
    ]
    _run(cmd, cwd=root)

    _print_header("PREPARANDO PACOTE FINAL")
    exe_path = _find_built_exe(dist_dir)
    dist_app_dir = dist_dir / APP_NAME
    dist_app_dir.mkdir(parents=True, exist_ok=True)

    final_exe = dist_app_dir / exe_path.name
    if final_exe.exists():
        final_exe.unlink()
    shutil.move(str(exe_path), str(final_exe))
    print(f"[OK] Executável movido para: {final_exe}")

    for file_name in COPY_FILES:
        _safe_copy_file(root / file_name, dist_app_dir)

    for dir_name in COPY_DIRS:
        _safe_copy_dir(root / dir_name, dist_app_dir)

    _ensure_runtime_dirs(dist_app_dir, RUNTIME_DIRS)
    _write_launcher_bat(dist_app_dir)
    _write_readme(dist_app_dir)

    _print_header("BUILD CONCLUÍDO")
    print(f"Pacote final: {dist_app_dir}")
    print(f"Executável  : {final_exe}")
    print("Para gerar novamente no futuro, execute: python build_exe.py")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as e:
        print()
        print("ERRO: falha ao executar um comando externo.")
        print(f"Código de saída: {e.returncode}")
        raise
    except Exception as e:
        print()
        print(f"ERRO: {e}")
        raise
