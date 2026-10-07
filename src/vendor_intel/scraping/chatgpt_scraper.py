"""ChatGPT web-UI scraper (vendored from github.com/st1vms/gptauto, MIT license;
ported from Firefox+selenium-wire to Chrome+plain Selenium during integration).

Unofficial: drives Google Chrome via Selenium to talk to chatgpt.com like a
human would. This is NOT the OpenAI API — prefer the ``openai`` client (see
``crawler/smart_crawl.py``) for anything programmatic; only reach for this
when a task specifically needs the ChatGPT web app itself.

Runs ANONYMOUS/guest by default (no login) — verified working end-to-end.
Google-account login does NOT reliably work here: Google's "this browser or
app may not be secure" check actively blocks Google OAuth from any
WebDriver (Selenium/Playwright/Puppeteer)-controlled browser, by design, with
no known workaround. If your OpenAI account only has Google SSO, anonymous
mode is the only option through this scraper. An OpenAI account with an
email+password login (not Google SSO) *may* work by logging into the Chrome
profile (``user_data_dir``) manually first — untested here since Google SSO
was the only login method available at integration time.

Requirements: none beyond what's already in requirements.txt — reuses the
same Selenium + Chrome install as vendor_intel.scraping.selenium_browser.

Known limitations (inherited from upstream, plus what we found integrating):
    - Headless mode is untested/likely unreliable for this specific flow —
      leave headless=False (the default) unless you've verified otherwise.
    - May hit captchas under heavy/rapid use.
    - May break on ChatGPT UI changes (CSS selectors are hardcoded; already
      had to be updated once from the original gptauto Firefox version).
    - Subject to OpenAI's terms for account automation — use at your own risk.

To reuse your own already-logged-in Chrome profile (e.g. to try a real
ChatGPT account instead of anonymous mode): close ALL Chrome windows first
(Chrome locks a profile to one process — Selenium's launch fails/opens a
blank profile otherwise), then pass profile_dir="%LOCALAPPDATA%\\Google\\
Chrome\\User Data" and profile_directory="Default" (or "Profile 1" / etc —
whichever folder matches the account, see chrome://version's "Profile Path"
or Local State's profile.info_cache for the mapping).

Env vars:
    CHATGPT_SCRAPER_ENABLED=true                    # opt-in gate
    CHATGPT_SCRAPER_CHROME_PROFILE_DIR=<path>        # Chrome's User Data root
    CHATGPT_SCRAPER_CHROME_PROFILE_DIRECTORY=<name>  # e.g. "Default", "Profile 1"
"""
from __future__ import annotations

import os
from time import perf_counter, sleep
from random import uniform
from uuid import UUID
from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, Iterator, Union

if TYPE_CHECKING:
    from selenium.webdriver.chrome.webdriver import WebDriver

CHATGPT_SCRAPER_ENABLED = False
CHROME_PROFILE_DIR: str = ""
CHROME_PROFILE_DIRECTORY: str = ""
# Copy the profile's login files to a private dir instead of driving the
# real profile in place — lets automation run while the user's own Chrome
# stays open, and dodges enterprise-policy restrictions on the managed
# profile. See _copy_profile_for_automation().
CHROME_PROFILE_COPY: bool = True

BASE_URL = "https://chatgpt.com/"

TEXTAREA_CSSS = 'textarea#mobile-composer-prompt, div[id="prompt-textarea"]'
# data-message-role (current UI) replaced the older data-message-author-role;
# keep both so this degrades gracefully across ChatGPT UI revisions.
CHAT_MESSAGE_CSSS = '[data-message-role="user"], [data-message-role="assistant"], [data-message-author-role]'
SEND_BUTTON_CSSS = 'button[aria-label="Send message"], button[data-testid="send-button"]'
# Present on the send button while ChatGPT is still streaming a reply; its
# absence (send button back to its idle icon) means generation has finished.
STOP_STREAMING_CSSS = 'button[data-testid="stop-button"]'
# Appears under a finished assistant turn — a fallback completion signal for
# replies fast enough that the streaming button is never actually observed.
COPY_RESPONSE_CSSS = 'button[aria-label="Copy response"], button[aria-label="Copy message"]'


@dataclass(frozen=True)
class ChatMessage:
    """message_id is a UUID when ChatGPT's DOM exposes one (older UI), or a
    stable positional id like "turn-0" (current UI has none) — treat it as an
    opaque per-conversation key, not necessarily a real UUID."""

    message_id: str
    text: str


@dataclass(frozen=True)
class AssistantMessage(ChatMessage):
    """Assistant chat message"""


@dataclass(frozen=True)
class UserMessage(ChatMessage):
    """User chat message"""


class DriverNotInitializedError(Exception):
    """Raised when a scraper operation is called before .start()"""


def random_sleep(_min: float, _max: float) -> None:
    sleep(round(uniform(_min, _max), 3))


def assert_uuid(text: str) -> str:
    return str(UUID(text, version=4))


def _strip_speaker_label(text: str) -> str:
    """Drop the leading screen-reader label ("You said:" / "ChatGPT said:")
    that .text picks up from the message container's accessibility markup."""
    first_line, _, rest = text.partition("\n")
    if first_line.strip().lower() in ("you said:", "chatgpt said:"):
        return rest.strip()
    return text.strip()


def _copy_profile_for_automation(
    user_data_dir: str, profile_directory: str | None
) -> str | None:
    """Clone the login-bearing parts of a real Chrome profile into a private
    directory and return that directory, or None if it can't be done.

    Why copy instead of using the profile directly:
      - Chrome locks a User Data directory to one process, so automating the
        real profile requires the user to close their browser entirely.
      - An enterprise-managed profile ("Managed by your organization") can
        refuse to expose a debugging port at all, which makes attaching
        impossible no matter what.
    A copy sidesteps both: it isn't locked, isn't policy-managed, and still
    carries the cookies that keep chatgpt.com logged in.

    Only the files needed for session/login are copied — not the whole
    profile, which can be gigabytes of cache and history.
    """
    import shutil
    import tempfile
    from pathlib import Path

    src_root = Path(user_data_dir)
    src_profile = src_root / (profile_directory or "Default")
    if not src_profile.is_dir():
        return None

    global _PROFILE_COPY_PATH
    if _PROFILE_COPY_PATH and Path(_PROFILE_COPY_PATH).is_dir():
        return _PROFILE_COPY_PATH  # reuse within a run; keeps the session warm

    try:
        dest_root = Path(tempfile.mkdtemp(prefix="chatgpt_scraper_profile_"))
        dest_profile = dest_root / "Default"
        dest_profile.mkdir(parents=True, exist_ok=True)

        # Cookies/tokens live in these; "Local State" holds the key used to
        # decrypt them, so it must come along or the cookies are unreadable.
        for name in ("Cookies", "Cookies-journal", "Login Data", "Web Data", "Preferences"):
            src = src_profile / name
            if src.is_file():
                shutil.copy2(src, dest_profile / name)
        local_state = src_root / "Local State"
        if local_state.is_file():
            shutil.copy2(local_state, dest_root / "Local State")

        _PROFILE_COPY_PATH = str(dest_root)
        return _PROFILE_COPY_PATH
    except Exception:
        # Any failure here is non-fatal: fall back to the profile as given.
        return None


_PROFILE_COPY_PATH: str = ""


class ChatGPTScraper:
    """Drives a Chrome session against chatgpt.com's web UI."""

    def __init__(
        self,
        profile_dir: str | None = None,
        profile_directory: str | None = None,
        headless: bool = False,
    ) -> None:
        self.driver: "WebDriver | None" = None
        self.messages: Dict[str, Union[UserMessage, AssistantMessage]] = {}
        self.headless = headless
        self.profile_dir = profile_dir or CHROME_PROFILE_DIR or None
        # e.g. "Default", "Profile 1", "Profile 2" — Chrome's User Data
        # subfolder for a specific account, distinct from --user-data-dir
        # (the User Data root itself).
        self.profile_directory = profile_directory or CHROME_PROFILE_DIRECTORY or None

    def __del__(self):
        self.quit()

    def _chrome_options(self):
        from vendor_intel.scraping.selenium_browser import build_chrome_options

        user_data_dir = self.profile_dir
        profile_directory = self.profile_directory
        if user_data_dir and CHROME_PROFILE_COPY:
            copied = _copy_profile_for_automation(user_data_dir, profile_directory)
            if copied:
                # The copy is a standalone User Data root holding just this
                # one profile, cloned into "Default".
                user_data_dir, profile_directory = copied, "Default"

        options = build_chrome_options(headless=self.headless, user_data_dir=user_data_dir)
        if profile_directory:
            options.add_argument(f"--profile-directory={profile_directory}")
        return options

    @staticmethod
    def assert_driver(func):
        def __wrapper(self: "ChatGPTScraper", *args, **kwargs):
            if self.driver is None or not self.driver.current_url.startswith(BASE_URL):
                raise DriverNotInitializedError(
                    "Initialize a web driver instance with .start() before performing actions"
                )
            return func(self, *args, **kwargs)

        return __wrapper

    def start(self, chat_id: str | None = None) -> None:
        try:
            from selenium import webdriver
        except ImportError as e:
            raise RuntimeError(
                "ChatGPTScraper requires selenium: pip install selenium "
                "(Google Chrome must also be installed)"
            ) from e

        if self.driver is not None:
            self.driver.quit()
            self.driver = None

        self.driver = webdriver.Chrome(options=self._chrome_options())

        url = BASE_URL
        if chat_id is not None:
            chat_id = assert_uuid(chat_id)
            url = f"{BASE_URL}/c/{chat_id}"

        self.driver.set_page_load_timeout(30)
        self.driver.get(url)
        self.driver.implicitly_wait(10)

    def quit(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.driver = None
        self.messages.clear()

    @assert_driver
    def current_chat_id(self) -> str | None:
        if not self.driver.current_url.startswith(f"{BASE_URL}/c/"):
            return None
        u = self.driver.current_url.split(f"{BASE_URL}/c/", maxsplit=1)[1]
        if not u:
            return None
        u = u.split("/", maxsplit=1)[0]
        return u or None

    @assert_driver
    def send_message(self, text: str, retries: int = 3) -> None:
        from selenium.common.exceptions import StaleElementReferenceException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        try:
            textarea = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, TEXTAREA_CSSS))
            )
            textarea.click()

            if textarea.tag_name.lower() == "textarea":
                # Real <textarea> (current ChatGPT UI) — type via real key
                # events so React's composer state (and the send button's
                # enabled state) updates the same way it would for a human
                # typing. Selenium's send_keys() treats a bare "\n" as an
                # Enter keypress (submits immediately), not a literal
                # newline — for any multi-line text (e.g. system+user
                # prompts joined by "\n\n") this silently truncated
                # everything after the first line. Split on newlines and
                # join with Shift+Enter (soft newline, does not submit) so
                # the full multi-line text lands in the composer intact.
                from selenium.webdriver.common.action_chains import ActionChains
                from selenium.webdriver.common.keys import Keys

                lines = text.split("\n")
                for i, line in enumerate(lines):
                    if line:
                        textarea.send_keys(line)
                    if i < len(lines) - 1:
                        ActionChains(self.driver).key_down(Keys.SHIFT).send_keys(
                            Keys.ENTER
                        ).key_up(Keys.SHIFT).perform()
            else:
                # Legacy contenteditable <div id="prompt-textarea">
                self.driver.execute_script(
                    "arguments[0].textContent=arguments[1]", textarea, text
                )
                textarea.send_keys(" ")  # nudge React to notice the DOM change

            random_sleep(1, 2)
            self.__click_send_button()
        except StaleElementReferenceException:
            # React re-rendered the composer between find and interact —
            # a transient race, not a real failure. Retry from scratch.
            if retries <= 0:
                raise
            random_sleep(0.5, 1)
            self.send_message(text, retries=retries - 1)

    def __click_send_button(self, retries: int = 3) -> None:
        from selenium.common.exceptions import StaleElementReferenceException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, SEND_BUTTON_CSSS))
        )
        try:
            button.click()
        except StaleElementReferenceException:
            if retries <= 0:
                raise
            self.__click_send_button(retries=retries - 1)

    @assert_driver
    def get_messages(self) -> Iterator[Union[UserMessage, AssistantMessage]]:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        elements = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_all_elements_located((By.CSS_SELECTOR, CHAT_MESSAGE_CSSS))
        )
        if not elements:
            return

        for index, div in enumerate(elements):
            role = div.get_attribute("data-message-role") or div.get_attribute(
                "data-message-author-role"
            )
            # Real UUID when the DOM exposes one (older UI); otherwise a
            # stable positional key — the current UI has no per-message id.
            msg_id = div.get_attribute("data-message-id") or f"turn-{index}"

            if msg_id in self.messages:
                yield self.messages[msg_id]
                continue

            text = _strip_speaker_label(div.text)
            if role == "user":
                self.messages[msg_id] = UserMessage(msg_id, text)
                yield self.messages[msg_id]
            elif role == "assistant":
                self.messages[msg_id] = AssistantMessage(msg_id, text)
                yield self.messages[msg_id]

    @assert_driver
    def wait_completion(self, timeout: float = 0) -> None:
        """Poll for the stop-streaming button to appear then disappear.

        Fast (e.g. anonymous/free-tier) replies can finish generating between
        two polls, so we may never observe the streaming button at all — a
        "Copy response" action on a finished assistant turn is an equally
        valid signal that generation is done."""
        from selenium.webdriver.common.by import By

        start = perf_counter()
        seen_streaming = False
        while timeout == 0 or perf_counter() - start <= timeout:
            streaming = bool(self.driver.find_elements(By.CSS_SELECTOR, STOP_STREAMING_CSSS))
            copy_response = bool(
                self.driver.find_elements(By.CSS_SELECTOR, COPY_RESPONSE_CSSS)
            )
            if streaming:
                seen_streaming = True
            elif seen_streaming or copy_response:
                random_sleep(0.5, 1)
                return
            sleep(0.25)
        if timeout != 0:
            raise TimeoutError(f"{timeout} seconds elapsed waiting for completion to finish")


def ask_chatgpt(
    text: str,
    *,
    profile_dir: str | None = None,
    profile_directory: str | None = None,
    timeout: float = 90,
    headless: bool = False,
) -> str:
    """One-shot helper: open a new chat, send `text`, return the assistant's reply."""
    if not CHATGPT_SCRAPER_ENABLED:
        raise RuntimeError("Set CHATGPT_SCRAPER_ENABLED=true to use ChatGPTScraper")

    scraper = ChatGPTScraper(
        profile_dir=profile_dir, profile_directory=profile_directory, headless=headless
    )
    try:
        scraper.start()
        scraper.send_message(text)
        scraper.wait_completion(timeout=timeout)
        messages = list(scraper.get_messages())
        if messages and isinstance(messages[-1], AssistantMessage):
            return messages[-1].text.strip()
        return ""
    finally:
        scraper.quit()


def apply_chatgpt_scraper_env() -> None:
    global CHATGPT_SCRAPER_ENABLED, CHROME_PROFILE_DIR, CHROME_PROFILE_DIRECTORY
    global CHROME_PROFILE_COPY

    raw = os.getenv("CHATGPT_SCRAPER_ENABLED", "").strip().lower()
    CHATGPT_SCRAPER_ENABLED = raw in ("1", "true", "yes", "on")

    if v := os.getenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIR", "").strip():
        CHROME_PROFILE_DIR = v
    if v := os.getenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIRECTORY", "").strip():
        CHROME_PROFILE_DIRECTORY = v

    raw_copy = os.getenv("CHATGPT_SCRAPER_CHROME_PROFILE_COPY", "").strip().lower()
    if raw_copy:
        CHROME_PROFILE_COPY = raw_copy in ("1", "true", "yes", "on")
