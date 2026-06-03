"""
Central configuration for ADMC Media Monitor.
Edit paths and defaults here.
"""

import os

# -- Base Directories
BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS_DIR  = os.path.join(BASE_DIR, "assets")
OUTPUT_DIR  = os.path.join(BASE_DIR, "output")

# -- Sub-directories
SNAPSHOTS_DIR = os.path.join(OUTPUT_DIR, "snapshots")
COVERAGE_DIR  = os.path.join(OUTPUT_DIR, "coverage")
TRACKER_DIR   = os.path.join(OUTPUT_DIR, "tracker")
LOGOS_DIR     = os.path.join(ASSETS_DIR, "client_logos")

# -- Logo Paths
ACTIVE_LOGO_PATH = os.path.join(ASSETS_DIR, "active_logo.png")

# -- File Naming Convention
# Snapshot  -> ADMC_<ClientName>_<MagazineName>_<DateOfPublishing>.docx
# Coverage  -> ADMC_<ClientName>_<PressReleaseName>_<Date>.docx
# Tracker   -> ADMC_clients_<ClientName>tracker.xlsx

SNAPSHOT_PREFIX = "ADMC"
COVERAGE_PREFIX = "ADMC"
TRACKER_PREFIX  = "ADMC_clients"

# -- Excel Tracker Columns
TRACKER_COLUMNS = [
    "S/N",
    "Date (DD/MM/YYYY)",
    "Publication",
    "Headline",
    "Media Type",
    "Language",
    "Spokesperson",
    "Source",
    "Reach",
    "Tier",
    "Print/Online",
    "For Pivot",
]

# -- 24h Coverage Columns
COVERAGE_COLUMNS = [
    "S/N",
    "Publication",
    "Headline",
    "Language",
    "Date",
    "Link",
    "Print/Online",
]

# -- Browser Settings
SCREENSHOT_WIDTH  = 1400
SCREENSHOT_HEIGHT = 900
PAGE_LOAD_WAIT    = 3          # seconds to wait after page load
SCROLL_PAUSE      = 0.5        # seconds between scrolls
FULL_PAGE         = True       # capture full page screenshot

# -- Ensure output directories exist
for _dir in [SNAPSHOTS_DIR, COVERAGE_DIR, TRACKER_DIR, LOGOS_DIR]:
    os.makedirs(_dir, exist_ok=True)
