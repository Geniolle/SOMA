from __future__ import annotations

import time

from soma_app.config.settings import Settings
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.webdriver_factory import WebDriverFactory


def main() -> int:
    settings = Settings.from_env()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    bundle = WebDriverFactory.create(settings)
    driver = bundle.driver
    try:
        driver.get("https://example.com")
        time.sleep(2)
    finally:
        driver.quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())