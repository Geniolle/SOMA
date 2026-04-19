from __future__ import annotations

import importlib
import sys
from pathlib import Path


def _ensure_src_path() -> None:
    base_dir = Path(__file__).resolve().parent
    src_dir = base_dir / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))


def main() -> int:
    _ensure_src_path()
    run_soma_main = importlib.import_module("soma_app.workflows.run_soma").main
    return int(run_soma_main())


if __name__ == "__main__":
    raise SystemExit(main())
