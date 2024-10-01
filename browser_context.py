from playwright.sync_api import sync_playwright


class BrowserContextManager:
    def __init__(self, browser_type="chromium"):
        self.browser_type = browser_type
        self.browser = None
        self.page = None

    def __enter__(self):
        self.playwright = sync_playwright().start()
        if self.browser_type == "chromium":
            self.browser = self.playwright.chromium.launch()
        elif self.browser_type == "firefox":
            self.browser = self.playwright.firefox.launch()
        elif self.browser_type == "webkit":
            self.browser = self.playwright.webkit.launch()
        else:
            raise ValueError("Invalid browser type")

        self.page = self.browser.new_page()
        return self.page

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.browser.close()
        self.playwright.stop()
