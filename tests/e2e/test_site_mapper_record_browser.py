from __future__ import annotations

import pytest

from soma_app.tools.record_browser_smoke_test import RecordBrowserSmokeTest, SmokeTestConfig

pytestmark = pytest.mark.e2e


def test_site_mapper_record_browser_smoke():
    result = RecordBrowserSmokeTest(
        SmokeTestConfig(
            prompt_timeout_seconds=180.0,
            record_active_timeout_seconds=90.0,
            command_timeout_seconds=60.0,
            exit_timeout_seconds=120.0,
        )
    ).run()

    assert result.status == "complete"
    assert result.returncode == 0
    assert result.artifacts_dir.exists()
    assert result.run_id.startswith("teste-browser-e2e-")
    assert "record_activated" in result.validations
    assert "artifacts_validated" in result.validations
