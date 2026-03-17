from __future__ import annotations

import logging

from soma_app.config.settings import Settings
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs


log = logging.getLogger("soma_app.main")


def main() -> int:
    settings = Settings.from_env()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    log.info(
        "SOMA boot OK | env=%s | headless=%s | timeout=%ss | retries=%s",
        settings.run_env,
        settings.headless,
        settings.timeout_seconds,
        settings.retry_count,
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())