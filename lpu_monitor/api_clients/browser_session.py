"""
browser_session.py

Manages a single Playwright browser context used to reach LPU's Cloudflare-protected
backend API (webapi.lpu.in) the same way a real browser does.

WHY THIS EXISTS
----------------
Plain HTTP libraries (requests/curl) are blocked by Cloudflare's bot-management
challenge on webapi.lpu.in ("Just a moment..." page instead of JSON). A real (or
"new-headless") Chromium instance, driven by Playwright, passes that challenge the
way a normal visitor would. Once the challenge is cleared, we don't reload a full
page for every programme -- instead we run fetch() *inside* that already-cleared
page's JS context for each subsequent call. That fetch carries the same cookies,
headers, and TLS fingerprint as a real browser request, because it genuinely is one.

This is deliberately NOT trying to defeat or reverse-engineer Cloudflare's challenge
logic itself -- it just lets a real browser do what real browsers are allowed to do,
which keeps this resilient to Cloudflare changing its internal detection rules later.
"""

import asyncio
import json
import logging
import random
from pathlib import Path
from typing import Any, Optional

from playwright.async_api import async_playwright, Browser, BrowserContext, Page

logger = logging.getLogger("lpu_monitor.browser_session")

# Text fragments that reliably show up on Cloudflare's interstitial/challenge page.
CHALLENGE_MARKERS = ("Just a moment", "cf-mitigated", "Checking your browser", "cf-chl-")


class CloudflareChallengeError(Exception):
    """Raised when a call returns Cloudflare's challenge page instead of the API's JSON."""


class BrowserSession:
    """
    Wraps one Playwright browser + context for the lifetime of a monitoring run.

    Usage:
        async with BrowserSession() as session:
            data = await session.call_api(
                url="https://webapi.lpu.in/webProgrammes/api/ProgramSearch/GetProgramFeeDetail",
                method="POST",
                body={"officialCode": "P132"},
            )
    """

    # Any real, working programme page. Used only to trigger and pass Cloudflare's
    # challenge naturally -- its content is otherwise irrelevant.
    ANCHOR_PAGE_URL = "https://www.lpu.in/programmes/engineering/b-tech-computer-science"

    def __init__(
        self,
        storage_state_path: Optional[str] = "data/session_state.json",
        headless: bool = True,
        min_delay_seconds: float = 1.0,
        max_delay_seconds: float = 2.5,
    ):
        self.storage_state_path = Path(storage_state_path) if storage_state_path else None
        self.headless = headless
        self.min_delay_seconds = min_delay_seconds
        self.max_delay_seconds = max_delay_seconds

        self._playwright = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None

    async def __aenter__(self) -> "BrowserSession":
        await self._start()
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        await self._stop()

    # ------------------------------------------------------------------ #
    # Lifecycle
    # ------------------------------------------------------------------ #

    async def _start(self) -> None:
        self._playwright = await async_playwright().start()

        # A real installed Chrome ("channel=chrome") is meaningfully harder for
        # Cloudflare to fingerprint as automation than Playwright's bundled
        # Chromium build. Falls back automatically if Chrome isn't installed on
        # this machine (e.g. a fresh CI runner) so setup never hard-fails.
        launch_kwargs: dict[str, Any] = {"headless": self.headless}
        try:
            self._browser = await self._playwright.chromium.launch(channel="chrome", **launch_kwargs)
        except Exception:
            logger.warning("Chrome channel not available, falling back to bundled Chromium")
            self._browser = await self._playwright.chromium.launch(**launch_kwargs)

        context_kwargs: dict[str, Any] = {
        "viewport": {"width": 1366, "height": 768},
        "locale": "en-IN",
        "ignore_https_errors": True,
        }
        # Reusing a saved session (cookies + localStorage) across runs means most
        # runs can skip the challenge-solving navigation entirely. If the saved
        # session has expired, _establish_session() below detects that and falls
        # back to a fresh navigation automatically -- so this is a pure speed
        # optimization, never a reliability risk.
        if self.storage_state_path and self.storage_state_path.exists():
            context_kwargs["storage_state"] = str(self.storage_state_path)
            logger.info("Reusing saved session state from %s", self.storage_state_path)

        self._context = await self._browser.new_context(**context_kwargs)
        self._page = await self._context.new_page()
        self._attach_diagnostic_listeners(self._page)

        await self._establish_session()

    def _attach_diagnostic_listeners(self, page: Page) -> None:
        """
        Registers Playwright-level (not JS-level) listeners so we can see the real
        reason a call fails, even when fetch() itself only reports a generic
        "TypeError: Failed to fetch" to our JS code.

        Specifically:
          - Chrome's DevTools console gets the full CORS rejection string (e.g.
            "...has been blocked by CORS policy: Response to preflight request
            doesn't pass access control check...") even though that string is
            deliberately withheld from the JS fetch() promise itself. Playwright
            can read the console the same way DevTools does.
          - page.on("requestfailed") exposes Chromium's internal net:: error code
            (e.g. net::ERR_FAILED for a CORS block vs a genuine connection error),
            which again never reaches JS.
          - page.on("response") logs every response that DID make it back,
            including its status and headers -- useful for confirming whether
            Cloudflare's challenge is being served instead of the real API
            response on a given host.
        """

        def on_console(msg) -> None:
            if msg.type in ("error", "warning"):
                logger.warning("[browser console %s] %s", msg.type, msg.text)

        def on_request_failed(request) -> None:
            if "webapi.lpu.in" in request.url:
                logger.warning(
                    "[requestfailed] %s %s -- failure: %s",
                    request.method, request.url, request.failure,
                )

        def on_response(response) -> None:
            if "webapi.lpu.in" in response.url:
                logger.info(
                    "[response] %s %s -> HTTP %s (headers: %s)",
                    response.request.method, response.url, response.status,
                    dict(response.headers),
                )

        page.on("console", on_console)
        page.on("requestfailed", on_request_failed)
        page.on("response", on_response)

    async def _stop(self) -> None:
        if self.storage_state_path and self._context:
            self.storage_state_path.parent.mkdir(parents=True, exist_ok=True)
            await self._context.storage_state(path=str(self.storage_state_path))
            logger.info("Saved session state to %s", self.storage_state_path)

        if self._context:
            await self._context.close()
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()

    async def _establish_session(self) -> None:
        """
        Navigate to a real programme page so Cloudflare's JS challenge resolves
        naturally, exactly as it would for a genuine visitor. This is what makes
        every subsequent in-page fetch() call trusted by Cloudflare.
        """
        logger.info("Establishing session via anchor page navigation")
        await self._page.goto(
            self.ANCHOR_PAGE_URL,
            wait_until="domcontentloaded",
            timeout=30000,
        )

        await asyncio.sleep(3)

        content = await self._page.content()
        if any(marker in content for marker in CHALLENGE_MARKERS):
            # Cloudflare's JS challenge often resolves itself within a few seconds
            # of the page finishing its own script execution. Give it a moment,
            # then check once more before treating this as a real failure.
            logger.info("Challenge page detected, waiting for automatic resolution")
            await asyncio.sleep(6)
            content = await self._page.content()
            if any(marker in content for marker in CHALLENGE_MARKERS):
                raise CloudflareChallengeError(
                    "Could not pass Cloudflare's challenge on the anchor page after "
                    "waiting. This may mean headless=False is required on this "
                    "network/IP, or that the IP itself needs to be allow-listed by "
                    "LPU IT (see the earlier design doc's Unknown #2/#3)."
                )
        logger.info("Session established successfully")

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    async def call_api(
        self,
        url: str,
        method: str = "POST",
        body: Optional[dict] = None,
        max_retries: int = 3,
        credentials_mode: str = "omit",
    ) -> Any:
        """
        Calls a webapi.lpu.in endpoint by running fetch() inside the already-
        challenge-cleared page context, so the request carries real browser
        cookies/headers/TLS fingerprint rather than a hand-built HTTP request.

        credentials_mode defaults to "omit" rather than "include": this endpoint
        looks like a stateless public data lookup (no login needed to see fee
        info), and "include" is the more likely cause of a CORS rejection if the
        server's Access-Control-Allow-Origin is a wildcard ("*"), since the fetch
        spec forbids combining a wildcard origin with credentialed requests. Pass
        credentials_mode="include" explicitly if you confirm from a DevTools
        capture that the real frontend does send cookies on this call.

        Returns the parsed JSON body on success. Raises RuntimeError if every
        retry attempt fails.
        """
        last_error: Optional[Exception] = None

        for attempt in range(1, max_retries + 1):
            # Small randomized delay between calls -- politeness towards LPU's
            # servers, and avoids a burst pattern that looks automated even to a
            # browser-based caller.
            await asyncio.sleep(random.uniform(self.min_delay_seconds, self.max_delay_seconds))
            try:
                result = await self._page.evaluate(
                    """
                    async ({ url, method, body, credentialsMode }) => {
                        try {
                            const resp = await fetch(url, {
                                method: method,
                                headers: { 'Content-Type': 'application/json' },
                                body: body ? JSON.stringify(body) : undefined,
                                credentials: credentialsMode,
                            });
                            const text = await resp.text();
                            return { status: resp.status, text: text, jsError: null };
                        } catch (err) {
                            // fetch() hides the real reason from JS, but capture whatever
                            // it *does* expose (name/message) -- the fuller reason will
                            // show up via the Playwright-level console/requestfailed
                            // listeners registered on this page instead.
                            return { status: null, text: null, jsError: `${err.name}: ${err.message}` };
                        }
                    }
                    """,
                    {"url": url, "method": method, "body": body, "credentialsMode": credentials_mode},
                )

                if result.get("jsError"):
                    raise RuntimeError(
                        f"fetch() rejected before any HTTP response for {url}: {result['jsError']} "
                        f"-- check the '[browser console]' and '[requestfailed]' log lines above "
                        f"for the actual CORS/network reason."
                    )

                if any(marker in result["text"] for marker in CHALLENGE_MARKERS):
                    raise CloudflareChallengeError(f"Challenge page returned for {url}")

                if result["status"] >= 400:
                    raise RuntimeError(
                        f"API returned HTTP {result['status']} for {url}: {result['text'][:200]}"
                    )

                return json.loads(result["text"])

            except CloudflareChallengeError as exc:
                last_error = exc
                logger.warning(
                    "Challenge encountered on attempt %d/%d for %s -- re-establishing session",
                    attempt, max_retries, url,
                )
                await self._establish_session()
            except Exception as exc:  # network errors, JSON decode errors, etc.
                last_error = exc
                logger.warning("Call failed on attempt %d/%d for %s: %s", attempt, max_retries, url, exc)
                await asyncio.sleep(2 ** attempt)

        raise RuntimeError(f"All {max_retries} attempts failed for {url}") from last_error

    async def call_api_via_request_context(
        self,
        url: str,
        method: str = "POST",
        body: Optional[dict] = None,
    ) -> Any:
        """
        DIAGNOSTIC / ALTERNATE PATH: calls the same endpoint using Playwright's
        own APIRequestContext (self._context.request) instead of an in-page
        fetch(). This is NOT executed inside the page's JS engine, so it is not
        subject to the browser's same-origin/CORS policy at all -- while still
        sharing the same cookie jar as the browser context (including any
        Cloudflare clearance cookie that IS valid for this exact host).

        If call_api() fails with "Failed to fetch" but THIS succeeds, that is
        strong confirmation the root cause is CORS specifically (hypothesis 1
        or 3), not a Cloudflare challenge on webapi.lpu.in itself (hypothesis 2)
        -- because a real Cloudflare challenge would block this path too.
        """
        request_fn = {
            "GET": self._context.request.get,
            "POST": self._context.request.post,
        }[method.upper()]

        response = await request_fn(
            url,
            data=json.dumps(body) if body else None,
            headers={"Content-Type": "application/json"},
        )
        text = await response.text()

        logger.info(
            "[request_context] %s %s -> HTTP %s (first 200 chars: %s)",
            method, url, response.status, text[:200],
        )

        if any(marker in text for marker in CHALLENGE_MARKERS):
            raise CloudflareChallengeError(f"Challenge page returned for {url} (via request context)")

        if response.status >= 400:
            raise RuntimeError(f"API returned HTTP {response.status} for {url}: {text[:200]}")

        return json.loads(text)
