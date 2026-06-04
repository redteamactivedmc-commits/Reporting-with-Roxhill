"""
Headless browser utilities for taking full-page screenshots.
"""

import time
import os
from PIL import Image

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from media_monitor.config import (
    SCREENSHOT_WIDTH,
    SCREENSHOT_HEIGHT,
    PAGE_LOAD_WAIT,
    SCROLL_PAUSE,
    FULL_PAGE,
)


def _build_driver() -> webdriver.Chrome:
    """Configure and return a headless Chrome driver."""
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument(f"--window-size={SCREENSHOT_WIDTH},{SCREENSHOT_HEIGHT}")
    options.add_argument("--hide-scrollbars")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/121.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)
    return driver


def _scroll_to_bottom(driver: webdriver.Chrome) -> None:
    """Scroll the page to ensure lazy-loaded content is rendered."""
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(SCROLL_PAUSE)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.3)


def capture_url_screenshot(url: str, output_path: str) -> str:
    """
    Navigate to `url`, take a full-page screenshot,
    save as PNG to `output_path`, and return the path.
    """
    driver = _build_driver()
    try:
        driver.get(url)
        time.sleep(PAGE_LOAD_WAIT)

        if FULL_PAGE:
            _scroll_to_bottom(driver)
            total_height = driver.execute_script("return document.body.scrollHeight")
            driver.set_window_size(SCREENSHOT_WIDTH, total_height)
            time.sleep(0.5)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        driver.save_screenshot(output_path)
        return output_path

    finally:
        driver.quit()


def screenshot_to_pil(path: str) -> Image.Image:
    """Load a screenshot PNG as a PIL Image."""
    return Image.open(path)
