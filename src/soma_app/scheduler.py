from __future__ import annotations

import os
import time
from datetime import datetime

from soma_app.workflows.run_soma import main as run_once


def _int_env(name: str, default: int) -> int:
    value = (os.getenv(name) or "").strip()
    try:
        parsed = int(value)
    except ValueError:
        return default
    return parsed if parsed > 0 else default


def _now() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def main() -> int:
    interval_seconds = _int_env("AGENDADOR_INTERVAL_SECONDS", 120)

    print("Agendador do SOMA iniciado.")
    print(f"Intervalo configurado: {interval_seconds} segundos.")

    while True:
        print("\n==================================================================")
        print(f"Iniciando nova execucao - {_now()}")
        print("==================================================================")

        try:
            exit_code = run_once()
            print(f"Execucao finalizada com exit_code={exit_code}.")
        except KeyboardInterrupt:
            print("\nAgendador interrompido pelo utilizador.")
            return 130
        except Exception as exc:
            print(f"Falha na execucao: {exc}")

        print(f"Aguardando {interval_seconds} segundos para a proxima ronda...")

        try:
            time.sleep(interval_seconds)
        except KeyboardInterrupt:
            print("\nAgendador interrompido pelo utilizador.")
            return 130
