# Metal Archives Scraper

Looks up releases from your music collection on [Metal Archives](https://www.metal-archives.com/) and writes the results back to your spreadsheet. Given a `.xlsx` file with Artist and Release columns, it searches for each release, fetches the matching band page, and fills in the Metal Archives URLs, release type, year, and genre. Ambiguous or unresolvable matches are flagged for manual review.

## How it works

The scraper runs as a visible Chromium browser window (via Playwright) so that it can maintain a real browser session with Cloudflare. It does not run headless. Each run picks artists at random from the unsearched rows in your spreadsheet and processes all of their releases before moving to the next artist. Progress is saved after each artist, so the script can be stopped and restarted at any point.

For each release it:

1. Searches the Metal Archives API (`/search/ajax-advanced/searching/albums`) by artist name and release title
2. If the exact-band search returns nothing, retries with fuzzy band matching
3. Fetches the band page for each search result and compares the full discography against your collection using fuzzy title matching
4. Picks the best-matching band and writes the results back to the spreadsheet

The rate limiter starts at configured intervals and automatically decreases them after sustained success, probing for headroom. It resets to the baseline after an error.

## Spreadsheet format

The scraper reads from and writes to a single `.xlsx` file. The first sheet is your collection. It must have at minimum:

| Column | Required | Description |
|--------|----------|-------------|
| Artist | Yes | Band name |
| Release | Yes | Album/EP title (`s/t` is expanded to the artist name automatically) |
| Year | No | Filled in by the scraper if empty |
| Genre | No | Filled in by the scraper if empty |
| Metal Archives Artist URL | No | Written by the scraper |
| Metal Archives Release URL | No | Written by the scraper |
| Type | No | Release type (Full-length, EP, etc.) — written by the scraper |
| Searched | No | Set to `True` once a release has been processed |
| Found | No | Set to `True` if a match was found, `False` otherwise |
| Needs Review | No | Set to `True` for ambiguous matches |

The scraper also creates (or appends to) three additional sheets:

- **Artists** — one row per unique Metal Archives band page, with name, location, and a JSON blob of band metadata
- **Not Found** — releases that returned no search results
- **Review** — releases that need human attention (partial fuzzy matches, multiple bands with the same name, releases that appear under multiple titles in the discography, etc.)

## Setup

**Requirements:** Python 3.9+, and a Playwright Chromium installation.

```bash
pip3 install -r requirements.txt
playwright install chromium
```

## Running

```bash
python3 main.py
```

On first run, you will be prompted to enter the path to your `.xlsx` file. The path is saved to `scraper_config.json` and reused on subsequent runs.

You can also pass the spreadsheet path directly:

```bash
python3 main.py --spreadsheet /path/to/collection.xlsx
```

When the browser opens, it navigates to Metal Archives to establish a session. If Cloudflare shows a challenge page, the script will pause and print a prompt — complete the challenge in the browser window and the script will continue automatically.

Press `Ctrl-C` at any time to stop. Progress up to the last completed artist is saved.

## Configuration

Optional settings can be set in `scraper_config.json` in the project root:

```json
{
  "query_interval_seconds": 15.0,
  "query_interval_min": 5.0,
  "artist_page_interval_seconds": 10.0,
  "artist_page_interval_min": 5.0,
  "fuzzy_match_threshold": 80
}
```

The rate limiter starts at `*_seconds` and adaptively decreases toward `*_min` after sustained successful requests. `fuzzy_match_threshold` (0–100) controls how closely a release title must match a Metal Archives discography entry to be considered a match. Lower values match more aggressively but increase false positives.

## Limitations

**Requires a visible browser.** The scraper cannot run headless. Cloudflare will block headless Chromium, so the browser window must remain open. The persistent browser profile (`browser_data/`) stores cookies and session state across runs, which reduces but does not eliminate Cloudflare challenges.

**Cloudflare challenges require manual intervention.** If Cloudflare presents a challenge mid-run, the scraper pauses for up to 5 minutes and prints an `[ACTION REQUIRED]` message. If the challenge is not completed in time, the script stops. The spreadsheet is saved at the end of each artist, so no progress is lost — the interrupted artist will be retried on the next run.

**Metal Archives coverage.** Only bands and releases indexed on Metal Archives are findable. Releases outside the site's genre scope (non-metal artists, unofficial releases, etc.) will land on the Not Found sheet regardless of how good the match logic is.

**Fuzzy matching has false positives and false negatives.** Release titles are compared using token-sort fuzzy matching. Titles that are very short, share many words with unrelated releases, or are transliterated differently may match incorrectly or fail to match. Any match flagged `Needs Review` should be checked manually.

**Single-threaded, single-browser.** Requests are made sequentially through one browser page. There is no parallelism. For large collections this is slow by design — Metal Archives rate-limits aggressively.

**No proxy or IP rotation support.** If Cloudflare blocks your IP entirely, the script will time out on every challenge. The only remedies are waiting, switching networks, or clearing the browser profile and re-establishing a session.
