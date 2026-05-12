import json
import os
from pathlib import Path

_PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_FILE = _PROJECT_ROOT / "scraper_config.json"
LOG_DIR = _PROJECT_ROOT / "logs"

def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}


def _cfg_float(key: str, default: float) -> float:
    return float(load_config().get(key, default))


# Intervals can be overridden in scraper_config.json, e.g.:
# {"query_interval_seconds": 12, "query_interval_min": 5}
QUERY_INTERVAL_SECONDS: float = _cfg_float("query_interval_seconds", 10.0)
QUERY_INTERVAL_MIN: float = _cfg_float("query_interval_min", 5.0)
ARTIST_PAGE_INTERVAL_SECONDS: float = _cfg_float("artist_page_interval_seconds", 8.0)
ARTIST_PAGE_INTERVAL_MIN: float = _cfg_float("artist_page_interval_min", 4.0)
FUZZY_MATCH_THRESHOLD: int = int(_cfg_float("fuzzy_match_threshold", 80))
# How many artists to process between periodic backups (0 = disabled)
BACKUP_INTERVAL: int = int(_cfg_float("backup_interval", 50))


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


def get_spreadsheet_path() -> str:
    cfg = load_config()
    path = cfg.get("spreadsheet_path")

    if path and os.path.isfile(path):
        return path

    if path and not os.path.isfile(path):
        print(f"Previously configured spreadsheet not found at: {path}")

    while True:
        entered = input("Enter the full path to your collection spreadsheet (.xlsx): ").strip()
        entered = entered.strip("'\"")
        if os.path.isfile(entered):
            cfg["spreadsheet_path"] = entered
            save_config(cfg)
            return entered
        print(f"File not found: {entered}. Please try again.")
