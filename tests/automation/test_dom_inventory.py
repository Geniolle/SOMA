from __future__ import annotations

from types import SimpleNamespace

from selenium.webdriver.common.by import By

from soma_app.automation.dom_inventory import (
    CaptureTracker,
    SelectorCandidate,
    absolute_xpath_from_segments,
    best_selector_candidate,
    build_absolute_xpath_from_segments,
    css_selector_for_element,
    is_dangerous_selector,
    is_dangerous_text,
    is_safe_auto_click,
    normalize_url,
    redact_sensitive_html,
    score_selector_candidate,
    selector_candidates_for_element,
    validate_selector,
    xpath_relative_for_element,
)


class FakeDriver:
    def __init__(self, counts: dict[tuple[str, str], int]):
        self.counts = counts

    def find_elements(self, by: str, selector: str):
        return [SimpleNamespace()] * self.counts.get((by, selector), 0)


def test_css_selector_prefers_data_testid():
    element = {
        "tag": "button",
        "id": "submit-1",
        "name": "submit",
        "class_name": "btn btn-primary",
        "data_attrs": {"data-testid": "login-submit"},
    }

    assert css_selector_for_element(element) == 'button[data-testid="login-submit"]'
    assert xpath_relative_for_element(element) == "//*[@data-testid='login-submit']"


def test_xpath_uses_label_association_for_inputs():
    element = {"tag": "input", "id": "", "name": "", "class_name": ""}
    assert xpath_relative_for_element(element, label_text="Descrição") == "//label[contains(normalize-space(.), 'Descrição')]/following::input[1]"


def test_absolute_xpath_generation():
    assert build_absolute_xpath_from_segments([("html", 1), ("body", 1), ("div", 2)]) == "/html[1]/body[1]/div[2]"
    assert absolute_xpath_from_segments([("html", 1), ("body", 1), ("div", 2)]) == "/html[1]/body[1]/div[2]"


def test_selector_validation_counts_and_best_selector():
    candidates = [
        SelectorCandidate("data-testid", By.CSS_SELECTOR, '[data-testid="login-submit"]', 100.0, "unique"),
        SelectorCandidate("absolute", By.XPATH, "/html/body/div[7]", 20.0, "fallback"),
    ]
    driver = FakeDriver({(By.CSS_SELECTOR, '[data-testid="login-submit"]'): 1, (By.XPATH, "/html/body/div[7]"): 3})

    validated = [validate_selector(driver, candidate) for candidate in candidates]
    assert validated[0].unique is True
    assert validated[0].count == 1
    assert validated[1].unique is False
    assert validated[1].count == 3

    best = best_selector_candidate(driver, candidates)
    assert best is not None
    assert best.selector == '[data-testid="login-submit"]'
    assert best.unique is True


def test_scoring_prefers_unique_data_testid_over_absolute_xpath():
    data_candidate = SelectorCandidate("data-testid", By.CSS_SELECTOR, '[data-testid="login-submit"]', 100.0, "unique")
    xpath_candidate = SelectorCandidate("absolute", By.XPATH, "/html/body/div[7]", 20.0, "fallback")

    data_score = score_selector_candidate(data_candidate, count=1)
    xpath_score = score_selector_candidate(xpath_candidate, count=3)

    assert data_score > xpath_score


def test_sanitization_redacts_sensitive_html():
    html = """
    <div>
      <input type="password" value="secret123">
      <input value="R$ 123,45">
      <span>user@example.com</span>
      <textarea>conteudo sensivel</textarea>
    </div>
    """

    sanitized = redact_sensitive_html(html)
    assert "secret123" not in sanitized
    assert "R$ 123,45" not in sanitized
    assert "user@example.com" not in sanitized
    assert "[redacted]" in sanitized


def test_dangerous_action_detection_blocks_submit_and_text():
    assert is_dangerous_text("Salvar documento") is True
    assert is_dangerous_text("Abrir cadastro") is False
    assert is_dangerous_selector("button[type='submit']") is True
    assert is_dangerous_selector(".btn-primary") is False
    assert is_safe_auto_click(tag="button", text="Salvar", selector="button", in_form=False) is False
    assert is_safe_auto_click(tag="button", text="Abrir", selector="button", in_form=True) is False


def test_normalize_url_strips_fragment_and_sorts_query():
    assert normalize_url("HTTPS://Example.com:443/a/?b=2&a=1#frag") == "https://example.com/a?a=1&b=2"


def test_capture_tracker_prevents_duplicates():
    tracker = CaptureTracker()
    assert tracker.should_capture("sig-1") is True
    assert tracker.should_capture("sig-1") is False
    tracker.remember("sig-2")
    assert tracker.should_capture("sig-2") is False


def test_selector_candidates_include_data_and_name():
    element = {
        "tag": "input",
        "id": "",
        "name": "email",
        "class_name": "form-control",
        "data_attrs": {"data-qa": "login-email"},
    }
    candidates = selector_candidates_for_element(element, label_text="Email")
    selectors = {candidate.selector for candidate in candidates}
    assert '[data-qa="login-email"]' in selectors
    assert '[name="email"]' in selectors or 'input[name="email"]' in selectors
