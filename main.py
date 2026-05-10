#!/usr/bin/env python3
from __future__ import annotations

import argparse
import logging
import random
import shutil
import time
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

from metal_archives_scraper import browser, scraper, spreadsheet
from metal_archives_scraper.config import (
    ARTIST_PAGE_INTERVAL_MIN,
    ARTIST_PAGE_INTERVAL_SECONDS,
    BACKUP_INTERVAL,
    LOG_DIR,
    QUERY_INTERVAL_MIN,
    QUERY_INTERVAL_SECONDS,
    get_spreadsheet_path,
)

LOG_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        RotatingFileHandler(LOG_DIR / "scraper.log", maxBytes=5_000_000, backupCount=3),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


_TYPE_PRECEDENCE = ["Full-length", "EP", "Single", "Live album", "Compilation", "Demo"]


def _save_periodic_backup(path: str) -> None:
    """Copy the current spreadsheet to a timestamped file in a backups/ subdirectory."""
    src = Path(path)
    backup_dir = src.parent / "backups"
    backup_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = backup_dir / f"{src.stem}_{timestamp}{src.suffix}"
    try:
        shutil.copy2(src, dst)
        logger.info("Periodic backup saved to %s", dst)
    except Exception as e:
        logger.warning("Periodic backup failed: %s", e)


def _pick_best_by_type(candidates: list) -> dict:
    """From same-titled discography entries, pick the one with the highest type precedence."""
    def _rank(entry):
        try:
            return _TYPE_PRECEDENCE.index(entry.get("type", ""))
        except ValueError:
            return len(_TYPE_PRECEDENCE)
    return min(candidates, key=_rank)


class AdaptiveThrottle:
    """Rate limiter that probes for headroom by shortening the interval after sustained
    successes, and resets to the baseline interval after an error."""

    def __init__(
        self,
        initial: float,
        min_interval: float,
        *,
        jitter: float = 0.25,
        decrease_factor: float = 0.85,
        successes_to_decrease: int = 10,
    ):
        self.interval = initial
        self._baseline = initial
        self._min = min_interval
        self._jitter = jitter
        self._decrease_factor = decrease_factor
        self._successes_to_decrease = successes_to_decrease
        self._consecutive_successes = 0
        self._last_time: float = 0.0

    def wait(self, label: str) -> None:
        elapsed = time.time() - self._last_time
        jitter_offset = self.interval * self._jitter * random.uniform(-1, 1)
        target = self.interval + jitter_offset
        wait_secs = target - elapsed
        if wait_secs > 0:
            logger.info(
                "Rate limit (%s): waiting %.1fs (interval=%.1fs)",
                label, wait_secs, self.interval,
            )
            time.sleep(wait_secs)

    def success(self) -> None:
        self._last_time = time.time()
        self._consecutive_successes += 1
        if self._consecutive_successes >= self._successes_to_decrease:
            new_interval = max(self._min, self.interval * self._decrease_factor)
            if new_interval < self.interval:
                logger.info(
                    "Rate limit adapting: %.1fs → %.1fs (floor=%.1fs)",
                    self.interval, new_interval, self._min,
                )
            self.interval = new_interval
            self._consecutive_successes = 0

    def reset(self) -> None:
        if self.interval < self._baseline:
            logger.info(
                "Rate limit reset to baseline %.1fs after error (was %.1fs)",
                self._baseline, self.interval,
            )
        self.interval = self._baseline
        self._consecutive_successes = 0
        self._last_time = time.time()


def _fetch_artist_pages(
    results: dict,
    artist_throttle: AdaptiveThrottle,
) -> tuple[list[dict], int]:
    """Fetch the artist info page for every URL returned by a search query."""
    artist_dicts: list[dict] = []
    fetch_errors = 0
    for artist_url in results:
        artist_throttle.wait("artist page")
        try:
            info = scraper.get_artist_info(artist_url)
            artist_dicts.append(info)
            artist_throttle.success()
        except RuntimeError:
            raise  # Cloudflare challenge timeout or fatal browser error — stop processing
        except Exception as e:
            logger.warning("Failed to get artist info for %s: %s", artist_url, e)
            fetch_errors += 1
    return artist_dicts, fetch_errors


def _process_artist(
    artist_name: str,
    release_titles: list[str],
    ws_collection,
    ws_artists,
    ws_review,
    ws_not_found,
    wb,
    path: str,
    query_throttle: AdaptiveThrottle,
    artist_throttle: AdaptiveThrottle,
) -> None:
    """Work through all releases for one artist."""
    release_titles = spreadsheet.expand_st_titles(
        ws_collection, ws_review, artist_name, release_titles
    )

    partial_match_titles: set[str] = set()

    for search_title in release_titles:

        # --- Search query ---
        query_throttle.wait("search query")
        logger.info("Searching: artist=%r release=%r", artist_name, search_title)
        results = scraper.search_artist_release(artist_name, search_title, exact_band=True)
        query_throttle.success()

        if not results:
            # Try fuzzy band match as fallback
            query_throttle.wait("search query (fuzzy)")
            logger.info("Exact match failed; retrying with exactBandMatch=0")
            results = scraper.search_artist_release(artist_name, search_title, exact_band=False)
            query_throttle.success()

        if not results:
            logger.info("No results for artist=%r release=%r", artist_name, search_title)
            spreadsheet.update_release_row(ws_collection, artist_name, search_title, searched=True, found=False)
            spreadsheet.add_not_found_entry(ws_not_found, artist_name, search_title)
            continue  # try next release title

        # --- Fetch artist pages ---
        artist_dicts, fetch_errors = _fetch_artist_pages(results, artist_throttle)

        if not artist_dicts:
            if fetch_errors > 0:
                # All fetches failed due to errors — leave rows unsearched for retry
                logger.warning(
                    "Skipping artist=%r release=%r due to fetch errors (will retry next run)",
                    artist_name, search_title,
                )
            else:
                spreadsheet.update_release_row(ws_collection, artist_name, search_title, searched=True, found=False)
                spreadsheet.add_not_found_entry(ws_not_found, artist_name, search_title)
            continue

        # --- Rank artists by release overlap ---
        ranked = scraper.rank_artists_by_releases(artist_dicts, release_titles)
        total_matches = sum(r["match_count"] for r in ranked)

        if total_matches == 0:
            # Search returned candidates but no release title cleared the fuzzy threshold.
            # Record as a partial match for human review.
            partial_urls = [
                e["release_url"] for e in results.values() if e.get("release_url")
            ]
            if partial_urls:
                spreadsheet.add_review_entry(
                    ws_review, artist_name, search_title,
                    "Partial match located",
                    partial_urls,
                )
            partial_match_titles.add(search_title)
            spreadsheet.update_release_row(
                ws_collection, artist_name, search_title, searched=True, found=False
            )
            continue

        # Exclude already-handled partial matches so _apply_artist_match
        # does not add them to the Not Found sheet.
        active_titles = [t for t in release_titles if t not in partial_match_titles]

        if len(ranked) == 1 or (len(ranked) > 1 and ranked[0]["match_count"] > 0 and ranked[1]["match_count"] == 0):
            # Unambiguous single best match
            _apply_artist_match(
                ranked[0], active_titles, artist_name,
                ws_collection, ws_artists, ws_review, ws_not_found,
            )
            break

        else:
            # Multiple artists with overlapping releases — flag for review
            all_matched_titles: set[str] = set()
            for r in ranked:
                if r["match_count"] > 0:
                    for m in r["matched_releases"]:
                        all_matched_titles.add(m["collection_title"])
                    _apply_artist_match(
                        r, active_titles, artist_name,
                        ws_collection, ws_artists, ws_review, ws_not_found,
                    )

            # Flag ambiguous releases on review sheet
            ambiguous_titles = [t for t in active_titles if t in all_matched_titles]
            if ambiguous_titles:
                urls = [r["artist"]["url"] for r in ranked if r["match_count"] > 0]
                for title in ambiguous_titles:
                    spreadsheet.add_review_entry(
                        ws_review, artist_name, title,
                        "Multiple Metal Archives artists matched this release",
                        urls,
                    )
                    spreadsheet.update_release_row(
                        ws_collection, artist_name, title,
                        searched=True, found=True, needs_review=True
                    )
            break

    spreadsheet.save_workbook(wb, path)


def _apply_artist_match(
    ranked_entry: dict,
    release_titles: list[str],
    artist_name: str,
    ws_collection,
    ws_artists,
    ws_review,
    ws_not_found,
):
    artist_dict = ranked_entry["artist"]
    matched = ranked_entry["matched_releases"]
    matched_collection_titles = {m["collection_title"] for m in matched}

    for m in matched:
        col_title = m["collection_title"]
        disco_matches_for_title = [
            d for d in artist_dict.get("discography", [])
            if d["title"] == m["discography_title"]
        ]
        if len(disco_matches_for_title) > 1:
            best = _pick_best_by_type(disco_matches_for_title)
            spreadsheet.update_release_row(
                ws_collection, artist_name, col_title,
                searched=True, found=True,
                artist_url=artist_dict["url"],
                release_url=best["url"],
                release_type=best["type"],
                year=best.get("year", ""),
                genre=artist_dict.get("genre", ""),
                needs_review=True,
            )
            spreadsheet.add_review_entry(
                ws_review, artist_name, col_title,
                "Multiple releases with this title exist in the discography",
                [d["url"] for d in disco_matches_for_title],
            )
        else:
            spreadsheet.update_release_row(
                ws_collection, artist_name, col_title,
                searched=True, found=True,
                artist_url=artist_dict["url"],
                release_url=m["discography_url"],
                release_type=m["discography_type"],
                year=m.get("discography_year", ""),
                genre=artist_dict.get("genre", ""),
            )

    # Mark unmatched releases as not found
    for title in release_titles:
        if title not in matched_collection_titles:
            spreadsheet.update_release_row(
                ws_collection, artist_name, title, searched=True, found=False
            )
            spreadsheet.add_not_found_entry(ws_not_found, artist_name, title)

    spreadsheet.update_artist_row(ws_artists, artist_dict)


def main():
    parser = argparse.ArgumentParser(description="Metal Archives Scraper")
    parser.add_argument(
        "--spreadsheet",
        type=str,
        default=None,
        help="Path to collection spreadsheet (overrides saved config).",
    )
    args = parser.parse_args()

    path = args.spreadsheet if args.spreadsheet else get_spreadsheet_path()
    logger.info("Using spreadsheet: %s", path)

    wb = spreadsheet.open_workbook(path)
    ws_collection = spreadsheet.ensure_collection_sheet(wb)
    ws_artists = spreadsheet.ensure_artists_sheet(wb)
    ws_review = spreadsheet.ensure_review_sheet(wb)
    ws_not_found = spreadsheet.ensure_not_found_sheet(wb)
    spreadsheet.save_workbook(wb, path)

    browser.launch_browser()
    logger.info("Starting main loop.")

    query_throttle = AdaptiveThrottle(QUERY_INTERVAL_SECONDS, QUERY_INTERVAL_MIN)
    artist_throttle = AdaptiveThrottle(ARTIST_PAGE_INTERVAL_SECONDS, ARTIST_PAGE_INTERVAL_MIN)

    artists_processed = 0

    try:
        while True:
            artist_name, release_titles = spreadsheet.pick_random_artist(ws_collection)
            if artist_name is None:
                logger.info("All releases have been searched. Exiting.")
                print("All releases have been searched.")
                break

            logger.info("Processing artist: %r (%d releases)", artist_name, len(release_titles))

            _process_artist(
                artist_name=artist_name,
                release_titles=release_titles,
                ws_collection=ws_collection,
                ws_artists=ws_artists,
                ws_review=ws_review,
                ws_not_found=ws_not_found,
                wb=wb,
                path=path,
                query_throttle=query_throttle,
                artist_throttle=artist_throttle,
            )

            artists_processed += 1
            if BACKUP_INTERVAL > 0 and artists_processed % BACKUP_INTERVAL == 0:
                _save_periodic_backup(path)

    except RuntimeError as e:
        logger.critical("Fatal browser error — stopping scraper: %s", e)
        print(f"\nScraper stopped due to a fatal browser error: {e}")
    except KeyboardInterrupt:
        logger.info("Interrupted by user. Progress has been saved.")
        print("\nScraper stopped. Progress has been saved to the spreadsheet.")
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
