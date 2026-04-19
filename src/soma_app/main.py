from __future__ import annotations

from soma_app.workflows.run_soma import main as run_soma_main


def main() -> int:
    return run_soma_main()


if __name__ == "__main__":
    raise SystemExit(main())
