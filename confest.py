import pytest

from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright, Page

 
@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as p:
        # Try to launch Chrome first, fall back to Chromium if Chrome is not available
        try:
            browser = p.chromium.launch(
                channel="chrome",
                headless=False,  # Set to True for container environment
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage"
                    "--disable-gpu",   # Disable GPU for container environments
                    "--window-size=1920,1080" # Set window size for consistency
                ]
            )
        except Exception as e:
            print(f"Failed to launch Chrome: {e}")
            print("Falling back to Chromium...")
            browser = p.chromium.launch(
                headless=True,  # Set to True for container environment
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",    # Disable GPU for container environments
                    "--window-size=1920,1080" # Set window size for consistency
                ]
            )
        yield browser
        browser.close()
 
@pytest.fixture
def page(browser, request):
    context = browser.new_context(
        no_viewport=True
    )
    page = context.new_page()
    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    yield page

    # Close page and context
    page.close()
    context.close()
 