#!/usr/bin/env python3
"""
Gerenciador de Logs do SOMA
Limpa logs antigos, compacta e rotaciona arquivos grandes
"""

import os
import sys
import gzip
import shutil
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Tuple


class LogManager:
    def __init__(self, log_dir: str = "logs", days_to_keep: int = 30, max_size_mb: int = 1024):
        self.log_dir = Path(log_dir)
        self.days_to_keep = days_to_keep
        self.max_size_bytes = max_size_mb * 1024 * 1024
        self.backup_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")

    def validate_dir(self) -> bool:
        """Validar se diretório de logs existe"""
        if not self.log_dir.exists():
            print(f"❌ Diretório {self.log_dir} não encontrado!")
            return False
        return True

    def get_directory_size(self) -> Tuple[int, str]:
        """Calcular tamanho total do diretório de logs"""
        total = sum(f.stat().st_size for f in self.log_dir.rglob("*") if f.is_file())
        return total, self._format_size(total)

    def find_old_logs(self) -> List[Path]:
        """Encontrar logs antigos (mais de X dias)"""
        cutoff = datetime.now() - timedelta(days=self.days_to_keep)
        old_logs = []

        for log_file in self.log_dir.glob("soma_dev_*.log"):
            mtime = datetime.fromtimestamp(log_file.stat().st_mtime)
            if mtime < cutoff:
                old_logs.append(log_file)

        return old_logs

    def delete_old_logs(self) -> int:
        """Deletar logs antigos"""
        old_logs = self.find_old_logs()

        if not old_logs:
            return 0

        for log_file in old_logs:
            try:
                log_file.unlink()
                print(f"  🗑️  Deletado: {log_file.name}")
            except Exception as e:
                print(f"  ⚠️  Erro ao deletar {log_file.name}: {e}")

        return len(old_logs)

    def find_empty_logs(self) -> List[Path]:
        """Encontrar arquivos de log vazios"""
        empty_logs = []

        for log_file in self.log_dir.glob("*.log"):
            if log_file.stat().st_size == 0:
                empty_logs.append(log_file)

        return empty_logs

    def delete_empty_logs(self) -> int:
        """Deletar logs vazios"""
        empty_logs = self.find_empty_logs()

        if not empty_logs:
            return 0

        for log_file in empty_logs:
            try:
                log_file.unlink()
                print(f"  🗑️  Deletado (vazio): {log_file.name}")
            except Exception as e:
                print(f"  ⚠️  Erro ao deletar {log_file.name}: {e}")

        return len(empty_logs)

    def rotate_main_log(self) -> bool:
        """Rotacionar arquivo principal soma-run.log se muito grande"""
        main_log = self.log_dir / "soma-run.log"

        if not main_log.exists():
            return False

        size = main_log.stat().st_size

        if size <= self.max_size_bytes:
            return False

        print(f"  ⚠️  soma-run.log muito grande: {self._format_size(size)}")
        print(f"  🔄 Rotacionando arquivo...")

        try:
            # Renomear arquivo
            backup_name = f"soma-run-{self.backup_suffix}.log"
            backup_path = self.log_dir / backup_name
            main_log.rename(backup_path)
            print(f"  ✅ Arquivo rotacionado para: {backup_name}")

            # Tentar compactação
            self._compress_log(backup_path)
            return True

        except Exception as e:
            print(f"  ❌ Erro ao rotacionar: {e}")
            return False

    def _compress_log(self, log_path: Path) -> bool:
        """Compactar arquivo de log"""
        try:
            gz_path = Path(str(log_path) + ".gz")

            with open(log_path, "rb") as f_in:
                with gzip.open(gz_path, "wb") as f_out:
                    shutil.copyfileobj(f_in, f_out)

            # Remover arquivo original
            log_path.unlink()
            print(f"  📦 Arquivo comprimido: {gz_path.name}")
            print(f"     Tamanho: {self._format_size(gz_path.stat().st_size)}")
            return True

        except Exception as e:
            print(f"  ⚠️  Não foi possível compactar: {e}")
            return False

    def list_logs(self, limit: int = 15) -> None:
        """Listar arquivos de log por tamanho"""
        logs = sorted(
            self.log_dir.glob("*"),
            key=lambda x: x.stat().st_size if x.is_file() else 0,
            reverse=True,
        )

        print(f"\n📋 ARQUIVOS DE LOG (primeiros {limit}):")
        print(f"{'Nome':<50} {'Tamanho':>12} {'Data':>15}")
        print("-" * 80)

        for log_file in logs[:limit]:
            if log_file.is_file():
                size = log_file.stat().st_size
                mtime = datetime.fromtimestamp(log_file.stat().st_mtime)
                print(f"{log_file.name:<50} {self._format_size(size):>12} {mtime.strftime('%Y-%m-%d %H:%M'):>15}")

    @staticmethod
    def _format_size(bytes_size: int) -> str:
        """Formatar tamanho em bytes para formato legível"""
        for unit in ["B", "KB", "MB", "GB"]:
            if bytes_size < 1024:
                return f"{bytes_size:.1f}{unit}"
            bytes_size /= 1024
        return f"{bytes_size:.1f}TB"

    def run(self) -> None:
        """Executar limpeza completa"""
        print("\n" + "=" * 80)
        print("GERENCIADOR DE LOGS - SOMA")
        print("=" * 80)

        if not self.validate_dir():
            sys.exit(1)

        # Tamanho antes
        size_before, size_before_str = self.get_directory_size()
        print(f"\n📊 ANTES DA LIMPEZA: {size_before_str}")

        # Limpeza
        print("\n🧹 EXECUTANDO LIMPEZA...")

        old_deleted = self.delete_old_logs()
        print(f"  ✅ {old_deleted} logs antigos removidos")

        empty_deleted = self.delete_empty_logs()
        print(f"  ✅ {empty_deleted} logs vazios removidos")

        rotated = self.rotate_main_log()
        if rotated:
            print(f"  ✅ Arquivo principal rotacionado")
        else:
            print(f"  ✅ Arquivo principal OK")

        # Tamanho depois
        size_after, size_after_str = self.get_directory_size()
        print(f"\n📊 DEPOIS DA LIMPEZA: {size_after_str}")

        freed = size_before - size_after
        print(f"💾 Espaço liberado: {self._format_size(freed)}")

        # Listar logs
        self.list_logs()

        print("\n" + "=" * 80)
        print("✅ LIMPEZA CONCLUÍDA!")
        print("=" * 80 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Gerenciador de logs do SOMA"
    )
    parser.add_argument(
        "-d", "--dir",
        default="logs",
        help="Diretório de logs (padrão: logs)"
    )
    parser.add_argument(
        "-k", "--keep-days",
        type=int,
        default=30,
        help="Manter logs dos últimos N dias (padrão: 30)"
    )
    parser.add_argument(
        "-m", "--max-size",
        type=int,
        default=1024,
        help="Tamanho máximo do arquivo principal em MB (padrão: 1024)"
    )

    args = parser.parse_args()

    manager = LogManager(
        log_dir=args.dir,
        days_to_keep=args.keep_days,
        max_size_mb=args.max_size
    )
    manager.run()


if __name__ == "__main__":
    main()
