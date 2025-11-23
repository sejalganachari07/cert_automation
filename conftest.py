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
        viewport={"width": 1920, "height": 1080}
    )

    context.tracing.start(screenshots=True, snapshots=True, sources=True)
    page = context.new_page()

    try:
        yield page
    finally:
        test_name = request.node.name
        trace_path = f"traces/{test_name}.zip"
        Path("traces").mkdir(exist_ok=True)

        context.tracing.stop(path=trace_path)   # MUST run before close
        page.close()
        context.close()

 