from __future__ import annotations

import sys
from typing import Sequence


def main(argv: Sequence[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] in {"audit-conciliacao", "--audit-conciliacao"}:
        from soma_app.workflows.audit_conciliacao import main as audit_main

        return audit_main()

    from soma_app.workflows.run_soma import main as run_main

    return run_main()
