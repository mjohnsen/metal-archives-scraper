import pytest
from unittest.mock import MagicMock

import metal_archives_scraper.browser as browser_module
from metal_archives_scraper.browser import (
    _is_challenge_page,
    _wait_for_challenge,
)


# ---------------------------------------------------------------------------
# _is_challenge_page
# ---------------------------------------------------------------------------

class TestIsChallengePage:
    @pytest.mark.parametrize("marker", [
        "cf-browser-verification",
        "cf-challenge",
        "Just a moment",
        "Enable JavaScript and cookies",
    ])
    def test_detects_each_cloudflare_marker(self, marker):
        assert _is_challenge_page(f"<html><body>{marker}</body></html>") is True

    def test_returns_false_for_normal_html(self):
        assert _is_challenge_page("<html><body>Welcome to Metal Archives</body></html>") is False

    def test_returns_false_for_empty_string(self):
        assert _is_challenge_page("") is False

    def test_marker_anywhere_in_content_triggers_true(self):
        long_html = "<html>" + ("x" * 500) + "Just a moment" + ("y" * 500) + "</html>"
        assert _is_challenge_page(long_html) is True


# ---------------------------------------------------------------------------
# _wait_for_challenge
# ---------------------------------------------------------------------------

class TestWaitForChallenge:
    def test_returns_when_challenge_clears(self, monkeypatch):
        page = MagicMock()
        page.content.side_effect = [
            "<html>cf-challenge still here</html>",
            "<html>Welcome to Metal Archives</html>",
        ]
        time_seq = iter([0.0, 1.0, 2.0, 3.0])
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(time_seq=time_seq, sleep_mock=MagicMock()))
        _wait_for_challenge(page, "http://example.com", timeout_seconds=60)

    def test_raises_runtime_error_on_timeout(self, monkeypatch):
        page = MagicMock()
        page.content.return_value = "<html>cf-challenge persists</html>"
        time_seq = iter([0.0, 200.0])
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(time_seq=time_seq, sleep_mock=MagicMock()))
        with pytest.raises(RuntimeError, match="Timed out"):
            _wait_for_challenge(page, "http://example.com", timeout_seconds=120)

    def test_sleeps_between_polls(self, monkeypatch):
        page = MagicMock()
        page.content.return_value = "<html>Welcome</html>"
        mock_sleep = MagicMock()
        time_seq = iter([0.0, 1.0])
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(time_seq=time_seq, sleep_mock=mock_sleep))
        _wait_for_challenge(page, "http://example.com", timeout_seconds=60)
        mock_sleep.assert_called_once_with(3)


# ---------------------------------------------------------------------------
# fetch_url
# ---------------------------------------------------------------------------

class TestFetchUrl:
    @pytest.fixture(autouse=True)
    def mock_page(self, monkeypatch):
        self.page = MagicMock()
        self.page.is_closed.return_value = False
        monkeypatch.setattr(browser_module, "_page", self.page)
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(sleep_mock=MagicMock()))

    def test_returns_page_content_on_success(self):
        self.page.goto.return_value = MagicMock(status=200)
        self.page.content.return_value = "<html>Good</html>"

        result = browser_module.fetch_url("http://example.com")

        assert result == "<html>Good</html>"

    def test_reuses_single_page_across_calls(self, monkeypatch):
        self.page.goto.return_value = MagicMock(status=200)
        self.page.content.return_value = "<html>Good</html>"

        mock_ctx = MagicMock()
        monkeypatch.setattr(browser_module, "_context", mock_ctx)

        browser_module.fetch_url("http://example.com")
        browser_module.fetch_url("http://example.com/page2")

        mock_ctx.new_page.assert_not_called()

    def test_retries_on_transient_network_error(self):
        ok_response = MagicMock(status=200)
        self.page.goto.side_effect = [Exception("Connection reset"), ok_response]
        self.page.content.return_value = "<html>Good</html>"

        result = browser_module.fetch_url("http://example.com", retries=2)

        assert result == "<html>Good</html>"
        assert self.page.goto.call_count == 2

    def test_raises_after_all_retries_exhausted(self):
        self.page.goto.side_effect = Exception("Always fails")

        with pytest.raises(Exception, match="Always fails"):
            browser_module.fetch_url("http://example.com", retries=2)

        assert self.page.goto.call_count == 2

    def test_runtime_error_is_not_retried(self):
        self.page.goto.return_value = MagicMock(status=200)
        self.page.content.return_value = "<html>cf-challenge</html>"

        with pytest.raises(RuntimeError, match="Still on challenge page"):
            browser_module.fetch_url("http://example.com", retries=3)

        assert self.page.goto.call_count == 1

    def test_403_challenge_pauses_for_manual_completion(self, monkeypatch):
        self.page.goto.return_value = MagicMock(status=403)
        self.page.content.side_effect = [
            "<html>cf-challenge</html>",  # inside the 403 branch
            "<html>Welcome</html>",       # after _wait_for_challenge returns
        ]
        mock_wait = MagicMock()
        monkeypatch.setattr(browser_module, "_wait_for_challenge", mock_wait)

        browser_module.fetch_url("http://example.com")

        mock_wait.assert_called_once()


# ---------------------------------------------------------------------------
# fetch_html_fragment
# ---------------------------------------------------------------------------

class TestFetchHtmlFragment:
    @pytest.fixture(autouse=True)
    def mock_page(self, monkeypatch):
        self.page = MagicMock()
        self.page.is_closed.return_value = False
        self.page.url = "https://www.metal-archives.com/"
        monkeypatch.setattr(browser_module, "_page", self.page)
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(sleep_mock=MagicMock()))

    def test_returns_html_text_from_evaluate(self):
        self.page.evaluate.return_value = "<table><tr><td>Album</td></tr></table>"

        result = browser_module.fetch_html_fragment("https://www.metal-archives.com/band/discography/id/1/tab/all")

        assert result == "<table><tr><td>Album</td></tr></table>"
        self.page.evaluate.assert_called_once()

    def test_navigates_to_ma_home_when_not_on_ma(self):
        self.page.url = "about:blank"
        self.page.evaluate.return_value = "<table></table>"

        browser_module.fetch_html_fragment("https://www.metal-archives.com/band/discography/id/1/tab/all")

        self.page.goto.assert_called_once()

    def test_skips_home_navigation_when_already_on_ma(self):
        self.page.url = "https://www.metal-archives.com/bands/Pharaoh/2801"
        self.page.evaluate.return_value = "<table></table>"

        browser_module.fetch_html_fragment("https://www.metal-archives.com/band/discography/id/1/tab/all")

        self.page.goto.assert_not_called()

    def test_retries_on_transient_error(self):
        self.page.evaluate.side_effect = [Exception("Timeout"), "<table></table>"]

        result = browser_module.fetch_html_fragment("https://www.metal-archives.com/api", retries=2)

        assert result == "<table></table>"
        assert self.page.evaluate.call_count == 2

    def test_raises_after_all_retries_exhausted(self):
        self.page.evaluate.side_effect = Exception("Always fails")

        with pytest.raises(Exception, match="Always fails"):
            browser_module.fetch_html_fragment("https://www.metal-archives.com/api", retries=2)


# ---------------------------------------------------------------------------
# fetch_json
# ---------------------------------------------------------------------------

class TestFetchJson:
    @pytest.fixture(autouse=True)
    def mock_page(self, monkeypatch):
        self.page = MagicMock()
        self.page.is_closed.return_value = False
        self.page.url = "https://www.metal-archives.com/"
        monkeypatch.setattr(browser_module, "_page", self.page)
        monkeypatch.setattr(browser_module, "time",
                            _fake_time_module(sleep_mock=MagicMock()))

    def test_returns_parsed_json_via_evaluate(self):
        self.page.evaluate.return_value = {"aaData": [["row1"]]}

        result = browser_module.fetch_json("https://www.metal-archives.com/api")

        assert result == {"aaData": [["row1"]]}
        self.page.evaluate.assert_called_once()

    def test_navigates_to_ma_home_when_on_different_origin(self):
        self.page.url = "about:blank"
        self.page.evaluate.return_value = {"aaData": []}

        browser_module.fetch_json("https://www.metal-archives.com/api")

        self.page.goto.assert_called_once()

    def test_skips_home_navigation_when_already_on_ma(self):
        self.page.url = "https://www.metal-archives.com/bands/Bolt_Thrower/1"
        self.page.evaluate.return_value = {"aaData": []}

        browser_module.fetch_json("https://www.metal-archives.com/api")

        self.page.goto.assert_not_called()

    def test_retries_on_transient_error(self):
        self.page.evaluate.side_effect = [Exception("Timeout"), {"ok": True}]

        result = browser_module.fetch_json("https://www.metal-archives.com/api", retries=2)

        assert result == {"ok": True}
        assert self.page.evaluate.call_count == 2

    def test_raises_after_all_retries_exhausted(self):
        self.page.evaluate.side_effect = Exception("Always fails")

        with pytest.raises(Exception, match="Always fails"):
            browser_module.fetch_json("https://www.metal-archives.com/api", retries=2)

    def test_does_not_retry_runtime_error(self):
        self.page.evaluate.side_effect = RuntimeError("Something fatal")

        with pytest.raises(RuntimeError, match="Something fatal"):
            browser_module.fetch_json("https://www.metal-archives.com/api", retries=3)

        assert self.page.evaluate.call_count == 1


# ---------------------------------------------------------------------------
# close_browser
# ---------------------------------------------------------------------------

class TestCloseBrowser:
    def test_closes_page_context_and_stops_playwright(self, monkeypatch):
        mock_page = MagicMock()
        mock_page.is_closed.return_value = False
        mock_ctx = MagicMock()
        mock_pw = MagicMock()
        monkeypatch.setattr(browser_module, "_page", mock_page)
        monkeypatch.setattr(browser_module, "_context", mock_ctx)
        monkeypatch.setattr(browser_module, "_playwright", mock_pw)

        browser_module.close_browser()

        mock_page.close.assert_called_once()
        mock_ctx.close.assert_called_once()
        mock_pw.stop.assert_called_once()

    def test_skips_page_close_when_already_closed(self, monkeypatch):
        mock_page = MagicMock()
        mock_page.is_closed.return_value = True
        monkeypatch.setattr(browser_module, "_page", mock_page)
        monkeypatch.setattr(browser_module, "_context", MagicMock())
        monkeypatch.setattr(browser_module, "_playwright", MagicMock())

        browser_module.close_browser()

        mock_page.close.assert_not_called()

    def test_does_not_raise_when_close_errors(self, monkeypatch):
        mock_page = MagicMock()
        mock_page.is_closed.return_value = False
        mock_page.close.side_effect = Exception("Already closed")
        monkeypatch.setattr(browser_module, "_page", mock_page)
        monkeypatch.setattr(browser_module, "_context", MagicMock())
        monkeypatch.setattr(browser_module, "_playwright", MagicMock())

        browser_module.close_browser()  # should not raise


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fake_time_module(time_seq=None, sleep_mock=None):
    import itertools

    class FakeTime:
        def __init__(self):
            self._seq = iter(time_seq) if time_seq else itertools.repeat(0.0)
            self.sleep = sleep_mock or MagicMock()

        def time(self):
            return next(self._seq)

    return FakeTime()
