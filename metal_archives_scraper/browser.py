from __future__ import annotations

import logging
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, BrowserContext, Page, Playwright

logger = logging.getLogger(__name__)

_PROJECT_ROOT = Path(__file__).parent.parent
_USER_DATA_DIR = str(_PROJECT_ROOT / "browser_data")

_MA_HOME = "https://www.metal-archives.com/"

_playwright: Playwright = None
_context: BrowserContext = None
_page: Page = None


def launch_browser() -> BrowserContext:
    global _playwright, _context, _page
    _playwright = sync_playwright().start()
    _context = _playwright.chromium.launch_persistent_context(
        user_data_dir=_USER_DATA_DIR,
        headless=False,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-restore-last-session",
        ],
        user_agent=(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    _page = _context.new_page()
    logger.info("Browser launched.")
    _warm_up()
    return _context


def _warm_up():
    """Navigate to Metal Archives home page to establish a Cloudflare session."""
    logger.info("Warming up browser session on Metal Archives...")
    try:
        _page.goto(_MA_HOME, wait_until="networkidle", timeout=60000)
        content = _page.content()
        if _is_challenge_page(content):
            _wait_for_challenge(_page, _MA_HOME, timeout_seconds=300)
        logger.info("Session warm-up complete.")
    except Exception as e:
        logger.warning("Session warm-up failed (non-fatal): %s", e)


def _get_page() -> Page:
    global _page, _context
    if _page is None or _page.is_closed():
        _page = _context.new_page()
    return _page


def _is_challenge_page(content: str) -> bool:
    markers = [
        "cf-browser-verification",
        "cf-challenge",
        "Just a moment",
        "Enable JavaScript and cookies",
        "Verifying you are human",
        "DDoS protection by Cloudflare",
    ]
    return any(m in content for m in markers)


def _wait_for_challenge(page: Page, url: str, timeout_seconds: int = 300):
    logger.warning(
        "Cloudflare challenge detected at %s. Waiting up to %ds for manual completion.",
        url, timeout_seconds,
    )
    print(f"\n[ACTION REQUIRED] Cloudflare challenge on: {url}")
    print(f"Complete the challenge in the browser window. You have {timeout_seconds} seconds.")
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        time.sleep(3)
        content = page.content()
        if not _is_challenge_page(content):
            logger.info("Cloudflare challenge passed.")
            return
    raise RuntimeError("Timed out waiting for Cloudflare challenge completion.")


def fetch_url(url: str, retries: int = 3) -> str:
    page = _get_page()
    for attempt in range(1, retries + 1):
        try:
            response = page.goto(url, wait_until="networkidle", timeout=30000)

            if response and response.status == 403:
                content = page.content()
                if _is_challenge_page(content):
                    _wait_for_challenge(page, url)

            content = page.content()
            if _is_challenge_page(content):
                raise RuntimeError(f"Still on challenge page after wait: {url}")
            return content

        except RuntimeError:
            raise
        except Exception as e:
            logger.warning("fetch_url attempt %d/%d failed for %s: %s", attempt, retries, url, e)
            if attempt < retries:
                time.sleep(30)
            else:
                raise


def fetch_html_fragment(url: str, retries: int = 3) -> str:
    """Fetch a URL via XHR from the browser context, returning raw HTML text."""
    page = _get_page()
    if not page.url.startswith("https://www.metal-archives.com"):
        try:
            page.goto(_MA_HOME, wait_until="networkidle", timeout=30000)
        except Exception as e:
            logger.warning("Could not navigate to MA home before fetch_html_fragment: %s", e)
    for attempt in range(1, retries + 1):
        try:
            return page.evaluate(
                """
                async (url) => {
                    const resp = await fetch(url, {
                        headers: { 'X-Requested-With': 'XMLHttpRequest' }
                    });
                    if (!resp.ok) throw new Error('HTTP ' + resp.status);
                    return resp.text();
                }
                """,
                url,
            )
        except RuntimeError:
            raise
        except Exception as e:
            logger.warning("fetch_html_fragment attempt %d/%d failed for %s: %s", attempt, retries, url, e)
            if attempt < retries:
                time.sleep(30)
            else:
                raise


def fetch_json(url: str, retries: int = 3) -> dict:
    page = _get_page()
    if not page.url.startswith("https://www.metal-archives.com"):
        try:
            page.goto(_MA_HOME, wait_until="networkidle", timeout=30000)
        except Exception as e:
            logger.warning("Could not navigate to MA home before fetch_json: %s", e)
    for attempt in range(1, retries + 1):
        try:
            result = page.evaluate(
                """
                async (url) => {
                    const resp = await fetch(url, {
                        headers: {
                            'Accept': 'application/json, text/javascript, */*; q=0.01',
                            'X-Requested-With': 'XMLHttpRequest'
                        }
                    });
                    if (!resp.ok) {
                        throw new Error('HTTP ' + resp.status);
                    }
                    return resp.json();
                }
                """,
                url,
            )
            return result
        except RuntimeError:
            raise
        except Exception as e:
            logger.warning("fetch_json attempt %d/%d failed for %s: %s", attempt, retries, url, e)
            if attempt < retries:
                time.sleep(30)
            else:
                raise


def close_browser():
    global _playwright, _context, _page
    if _page and not _page.is_closed():
        try:
            # Navigate away so Chromium can flush its session state cleanly.
            _page.goto("about:blank", timeout=5000)
        except Exception:
            pass
        try:
            _page.close()
        except Exception as e:
            logger.warning("Error closing browser page: %s", e)
    if _context:
        try:
            _context.close()
        except Exception as e:
            logger.warning("Error closing browser context: %s", e)
    if _playwright:
        try:
            _playwright.stop()
        except Exception as e:
            logger.warning("Error stopping Playwright: %s", e)
