from __future__ import annotations

import logging
import re
from urllib.parse import urlencode

from bs4 import BeautifulSoup
from thefuzz import fuzz

from . import browser
from .config import FUZZY_MATCH_THRESHOLD

logger = logging.getLogger(__name__)

_SEARCH_AJAX = "https://www.metal-archives.com/search/ajax-advanced/searching/albums/"


def _normalize_years_active(val: str) -> str:
    """Collapse whitespace runs to single spaces and fix '(asFoo)' → '(as Foo)'."""
    val = re.sub(r"\s+", " ", val).strip()
    val = re.sub(r"\(as(?=[A-Za-z])", "(as ", val)
    return val
_DISCO_URL = "https://www.metal-archives.com/band/discography/id/{band_id}/tab/all"


def _parse_discography_html(html: str) -> list:
    soup = BeautifulSoup(html, "lxml")
    entries = []
    for tr in soup.select("tbody tr"):
        cells = tr.find_all("td")
        if not cells:
            continue
        title, rel_url = _extract_link(str(cells[0]))
        rel_type = cells[1].get_text(strip=True) if len(cells) > 1 else ""
        year = cells[2].get_text(strip=True) if len(cells) > 2 else ""
        entries.append({"title": title, "url": rel_url, "type": rel_type, "year": year})
    return entries


def _extract_link(html_fragment: str) -> tuple[str, str]:
    soup = BeautifulSoup(html_fragment, "lxml")
    tag = soup.find("a")
    if tag:
        return tag.get_text(strip=True), tag.get("href", "")
    return html_fragment, ""


def search_artist_release(artist_name: str, release_title: str, exact_band: bool = True) -> dict:
    params = {
        "bandName": artist_name,
        "exactBandMatch": "1" if exact_band else "0",
        "releaseTitle": release_title,
        "exactReleaseMatch": "0",
        "sEcho": "1",
        "iDisplayStart": "0",
        "iDisplayLength": "200",
    }
    url = _SEARCH_AJAX + "?" + urlencode(params)
    logger.info("Searching: artist=%r release=%r exact_band=%s", artist_name, release_title, exact_band)

    data = browser.fetch_json(url)
    results = {}

    for row in data.get("aaData", []):
        band_name, band_url = _extract_link(row[0])
        release_name, release_url = _extract_link(row[1])
        release_type = row[2] if len(row) > 2 else ""

        entry = {
            "artist_name": band_name,
            "artist_url": band_url,
            "release_title": release_name,
            "release_url": release_url,
            "release_type": release_type,
        }

        if band_url and band_url not in results:
            results[band_url] = entry

    return results


def get_artist_info(artist_url: str) -> dict:
    logger.info("Fetching artist page: %s", artist_url)
    html = browser.fetch_url(artist_url)
    soup = BeautifulSoup(html, "lxml")

    info: dict = {"url": artist_url}

    # Band name from <h1 class="band_name">
    name_tag = soup.find("h1", class_="band_name")
    info["name"] = name_tag.get_text(strip=True) if name_tag else ""

    # Band stats in <div id="band_stats"> which contains two <dl> elements
    stats = {}
    stats_div = soup.find("div", id="band_stats")
    if stats_div:
        for dl in stats_div.find_all("dl"):
            keys = [dt.get_text(strip=True).rstrip(":") for dt in dl.find_all("dt")]
            vals = [dd.get_text(strip=True) for dd in dl.find_all("dd")]
            stats.update(dict(zip(keys, vals)))

    info["country"] = stats.get("Country of origin", "")
    info["location"] = stats.get("Location", "")
    info["status"] = stats.get("Status", "")
    info["formed_in"] = stats.get("Formed in", "")
    info["genre"] = stats.get("Genre", "")
    info["themes"] = stats.get("Lyrical themes", "")
    info["label"] = stats.get("Current label", "")
    info["years_active"] = _normalize_years_active(stats.get("Years active", ""))

    # Extract numeric band ID from URL
    match = re.search(r"/(\d+)$", artist_url)
    band_id = match.group(1) if match else ""
    info["band_id"] = band_id

    # Fetch complete discography
    info["discography"] = []
    if band_id:
        disco_url = _DISCO_URL.format(band_id=band_id)
        logger.info("Fetching discography: %s", disco_url)
        try:
            html = browser.fetch_html_fragment(disco_url)
            info["discography"] = _parse_discography_html(html)
        except Exception as e:
            logger.warning("Failed to fetch discography for %s: %s", artist_url, e)

    return info


def rank_artists_by_releases(artist_dicts: list, release_titles: list) -> list:
    ranked = []

    for artist in artist_dicts:
        disco = artist.get("discography", [])
        matched = []
        used_disco_indices = set()

        for col_title in release_titles:
            best_score = 0
            best_match = None
            best_idx = None

            for idx, rel in enumerate(disco):
                if idx in used_disco_indices:
                    continue
                score = fuzz.token_sort_ratio(col_title, rel["title"])
                if score > best_score:
                    best_score = score
                    best_match = rel
                    best_idx = idx

            if best_score >= FUZZY_MATCH_THRESHOLD and best_match is not None:
                matched.append({
                    "collection_title": col_title,
                    "discography_title": best_match["title"],
                    "discography_url": best_match["url"],
                    "discography_type": best_match["type"],
                    "discography_year": best_match["year"],
                    "match_score": best_score,
                })
                used_disco_indices.add(best_idx)

        ranked.append({
            "artist": artist,
            "matched_releases": matched,
            "match_count": len(matched),
            "avg_score": (
                sum(m["match_score"] for m in matched) / len(matched) if matched else 0
            ),
        })

    ranked.sort(key=lambda x: (x["match_count"], x["avg_score"]), reverse=True)
    return ranked
