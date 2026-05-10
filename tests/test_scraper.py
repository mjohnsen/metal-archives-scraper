import pytest
from unittest.mock import MagicMock, patch

import metal_archives_scraper.browser as browser_module
from metal_archives_scraper.scraper import (
    _extract_link,
    _normalize_years_active,
    get_artist_info,
    rank_artists_by_releases,
    search_artist_release,
)

# ---------------------------------------------------------------------------
# _normalize_years_active
# ---------------------------------------------------------------------------

class TestNormalizeYearsActive:
    def test_collapses_tabs_and_newlines_to_single_space(self):
        raw = "1983\t\t\t(as VoltAge),\t\t\t\n\t\t\t1983-1985"
        assert _normalize_years_active(raw) == "1983 (as VoltAge), 1983-1985"

    def test_strips_leading_and_trailing_whitespace(self):
        assert _normalize_years_active("  1990-present  ") == "1990-present"

    def test_fixes_as_prefix_without_space(self):
        assert _normalize_years_active("(asVoltAge)") == "(as VoltAge)"

    def test_fixes_multiple_as_prefixes(self):
        result = _normalize_years_active("1980 (asOldName), 1985 (asOtherName)")
        assert "(as OldName)" in result
        assert "(as OtherName)" in result

    def test_does_not_alter_already_correct_as_prefix(self):
        assert _normalize_years_active("(as VoltAge)") == "(as VoltAge)"

    def test_returns_empty_string_unchanged(self):
        assert _normalize_years_active("") == ""

    def test_full_coroner_example(self):
        raw = "1983\t\t        \t\t\t(asVoltAge),\t\t        \t\t\t1983-1985,\t\t        \t\t\t1985-1996,\t\t        \t\t\t2010-present"
        result = _normalize_years_active(raw)
        assert result == "1983 (as VoltAge), 1983-1985, 1985-1996, 2010-present"


# ---------------------------------------------------------------------------
# Helpers / shared HTML fixtures
# ---------------------------------------------------------------------------

PHARAOH_BAND_HTML = """
<html><body>
  <h1 class="band_name">Pharaoh</h1>
  <div id="band_stats">
    <dl class="float_left">
      <dt>Country of origin:</dt><dd>United States</dd>
      <dt>Location:</dt><dd>Philadelphia, PA</dd>
      <dt>Status:</dt><dd>Active</dd>
      <dt>Formed in:</dt><dd>2000</dd>
    </dl>
    <dl class="float_right">
      <dt>Genre:</dt><dd>Power/Progressive Metal</dd>
      <dt>Lyrical themes:</dt><dd>Philosophy</dd>
      <dt>Current label:</dt><dd>Cruz Del Sur Music</dd>
      <dt>Years active:</dt><dd>2000-present</dd>
    </dl>
  </div>
</body></html>
"""

DISCO_HTML = """
<table width="100%">
<thead><tr><th>Name</th><th>Type</th><th>Year</th><th>Reviews</th></tr></thead>
<tbody>
<tr>
  <td><a href="https://www.metal-archives.com/albums/Pharaoh/After_the_Fire/123">After the Fire</a></td>
  <td>Full-length</td><td>2008</td><td>0</td>
</tr>
<tr>
  <td><a href="https://www.metal-archives.com/albums/Pharaoh/The_Longest_Night/456">The Longest Night</a></td>
  <td>Full-length</td><td>2006</td><td>2</td>
</tr>
</tbody>
</table>
"""

SEARCH_RESPONSE = {
    "aaData": [
        [
            "<a href='https://www.metal-archives.com/bands/Pharaoh/2801'>Pharaoh</a>",
            "<a href='https://www.metal-archives.com/albums/Pharaoh/After_the_Fire/123'>After the Fire</a>",
            "Full-length",
            "0",
        ]
    ]
}


# ---------------------------------------------------------------------------
# _extract_link
# ---------------------------------------------------------------------------

class TestExtractLink:
    def test_extracts_text_and_href_from_anchor(self):
        html = "<a href='https://example.com/bands/Foo/1'>Foo</a>"
        text, url = _extract_link(html)
        assert text == "Foo"
        assert url == "https://example.com/bands/Foo/1"

    def test_returns_raw_text_and_empty_url_when_no_anchor(self):
        text, url = _extract_link("Plain text")
        assert text == "Plain text"
        assert url == ""

    def test_handles_anchor_with_no_href(self):
        text, url = _extract_link("<a>No href here</a>")
        assert text == "No href here"
        assert url == ""

    def test_strips_whitespace_from_text(self):
        text, _ = _extract_link("<a href='x'>  Foo  </a>")
        assert text == "Foo"


# ---------------------------------------------------------------------------
# search_artist_release
# ---------------------------------------------------------------------------

class TestSearchArtistRelease:
    @pytest.fixture(autouse=True)
    def mock_fetch_json(self, monkeypatch):
        self.fetch_json = MagicMock(return_value=SEARCH_RESPONSE)
        monkeypatch.setattr(browser_module, "fetch_json", self.fetch_json)

    def test_returns_dict_keyed_by_artist_url(self):
        results = search_artist_release("Pharaoh", "After the Fire")
        assert "https://www.metal-archives.com/bands/Pharaoh/2801" in results

    def test_result_entry_contains_expected_fields(self):
        results = search_artist_release("Pharaoh", "After the Fire")
        entry = results["https://www.metal-archives.com/bands/Pharaoh/2801"]
        assert entry["artist_name"] == "Pharaoh"
        assert entry["release_title"] == "After the Fire"
        assert entry["release_type"] == "Full-length"
        assert "albums/Pharaoh/After_the_Fire/123" in entry["release_url"]

    def test_returns_empty_dict_when_no_results(self, monkeypatch):
        monkeypatch.setattr(browser_module, "fetch_json",
                            MagicMock(return_value={"aaData": []}))
        results = search_artist_release("Nonexistent", "Nothing")
        assert results == {}

    def test_deduplicates_same_artist_url(self, monkeypatch):
        two_rows_same_url = {
            "aaData": [
                [
                    "<a href='https://www.metal-archives.com/bands/Pharaoh/2801'>Pharaoh</a>",
                    "<a href='https://www.metal-archives.com/albums/Pharaoh/After_the_Fire/123'>After the Fire</a>",
                    "Full-length", "0",
                ],
                [
                    "<a href='https://www.metal-archives.com/bands/Pharaoh/2801'>Pharaoh</a>",
                    "<a href='https://www.metal-archives.com/albums/Pharaoh/Be_Gone/789'>Be Gone</a>",
                    "Full-length", "0",
                ],
            ]
        }
        monkeypatch.setattr(browser_module, "fetch_json", MagicMock(return_value=two_rows_same_url))
        results = search_artist_release("Pharaoh", "After the Fire")
        # Same URL → only one entry, first one wins
        assert len(results) == 1

    def test_uses_exact_band_match_by_default(self):
        search_artist_release("Pharaoh", "After the Fire")
        called_url = self.fetch_json.call_args[0][0]
        assert "exactBandMatch=1" in called_url

    def test_uses_inexact_band_match_when_specified(self):
        search_artist_release("Pharaoh", "After the Fire", exact_band=False)
        called_url = self.fetch_json.call_args[0][0]
        assert "exactBandMatch=0" in called_url

    def test_multiple_artists_with_different_urls_all_returned(self, monkeypatch):
        two_artists = {
            "aaData": [
                [
                    "<a href='https://www.metal-archives.com/bands/Pharaoh/2801'>Pharaoh</a>",
                    "<a href='https://ma.com/albums/A/1'>Album A</a>",
                    "Full-length", "0",
                ],
                [
                    "<a href='https://www.metal-archives.com/bands/Pharaoh/9999'>Pharaoh</a>",
                    "<a href='https://ma.com/albums/B/2'>Album B</a>",
                    "Full-length", "0",
                ],
            ]
        }
        monkeypatch.setattr(browser_module, "fetch_json", MagicMock(return_value=two_artists))
        results = search_artist_release("Pharaoh", "Album A")
        assert len(results) == 2


# ---------------------------------------------------------------------------
# get_artist_info
# ---------------------------------------------------------------------------

class TestGetArtistInfo:
    @pytest.fixture(autouse=True)
    def mock_browser(self, monkeypatch):
        monkeypatch.setattr(browser_module, "fetch_url", MagicMock(return_value=PHARAOH_BAND_HTML))
        monkeypatch.setattr(browser_module, "fetch_html_fragment", MagicMock(return_value=DISCO_HTML))

    def test_returns_band_name(self):
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        assert info["name"] == "Pharaoh"

    def test_returns_band_stats(self):
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        assert info["country"] == "United States"
        assert info["location"] == "Philadelphia, PA"
        assert info["status"] == "Active"
        assert info["formed_in"] == "2000"
        assert info["genre"] == "Power/Progressive Metal"
        assert info["themes"] == "Philosophy"
        assert info["label"] == "Cruz Del Sur Music"
        assert info["years_active"] == "2000-present"

    def test_extracts_band_id_from_url(self):
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        assert info["band_id"] == "2801"

    def test_returns_discography_list(self):
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        assert len(info["discography"]) == 2
        titles = [r["title"] for r in info["discography"]]
        assert "After the Fire" in titles
        assert "The Longest Night" in titles

    def test_discography_entry_has_all_fields(self):
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        first = info["discography"][0]
        assert "title" in first
        assert "url" in first
        assert "type" in first
        assert "year" in first

    def test_url_stored_in_result(self):
        url = "https://www.metal-archives.com/bands/Pharaoh/2801"
        info = get_artist_info(url)
        assert info["url"] == url

    def test_discography_empty_when_no_band_id(self, monkeypatch):
        no_id_html = "<html><body><h1 class='band_name'>Ghost Band</h1></body></html>"
        monkeypatch.setattr(browser_module, "fetch_url", MagicMock(return_value=no_id_html))
        info = get_artist_info("https://www.metal-archives.com/bands/GhostBand/")
        assert info["discography"] == []

    def test_discography_empty_on_fetch_failure(self, monkeypatch):
        monkeypatch.setattr(browser_module, "fetch_url", MagicMock(return_value=PHARAOH_BAND_HTML))
        monkeypatch.setattr(browser_module, "fetch_html_fragment",
                            MagicMock(side_effect=Exception("Network error")))
        info = get_artist_info("https://www.metal-archives.com/bands/Pharaoh/2801")
        assert info["discography"] == []

    def test_missing_band_stats_div_returns_empty_strings(self, monkeypatch):
        bare_html = "<html><body><h1 class='band_name'>NoStats</h1></body></html>"
        monkeypatch.setattr(browser_module, "fetch_url", MagicMock(return_value=bare_html))
        monkeypatch.setattr(browser_module, "fetch_html_fragment",
                            MagicMock(return_value="<table><tbody></tbody></table>"))
        info = get_artist_info("https://www.metal-archives.com/bands/NoStats/1")
        assert info["genre"] == ""
        assert info["country"] == ""


# ---------------------------------------------------------------------------
# rank_artists_by_releases
# ---------------------------------------------------------------------------

class TestRankArtistsByReleases:
    def _make_artist(self, name, titles):
        return {
            "name": name,
            "url": f"https://ma.com/bands/{name}/1",
            "discography": [
                {"title": t, "url": f"https://ma.com/albums/{name}/{t}/1",
                 "type": "Full-length", "year": "2000"}
                for t in titles
            ],
        }

    def test_artist_with_more_matches_ranks_first(self):
        a1 = self._make_artist("Pharaoh", ["After the Fire", "The Longest Night", "Be Gone"])
        a2 = self._make_artist("Pharaoh_2", ["Some Album"])
        collection = ["After the Fire", "The Longest Night"]

        ranked = rank_artists_by_releases([a1, a2], collection)

        assert ranked[0]["artist"]["name"] == "Pharaoh"
        assert ranked[0]["match_count"] == 2

    def test_fuzzy_matching_handles_minor_title_differences(self):
        artist = self._make_artist("Band", ["After the Fire"])
        collection = ["After The Fire"]  # different capitalisation

        ranked = rank_artists_by_releases([artist], collection)

        assert ranked[0]["match_count"] == 1
        assert ranked[0]["matched_releases"][0]["collection_title"] == "After The Fire"

    def test_word_order_variation_matched_by_token_sort_ratio(self):
        artist = self._make_artist("Band", ["Fire After the"])
        collection = ["After the Fire"]

        ranked = rank_artists_by_releases([artist], collection)

        assert ranked[0]["match_count"] == 1

    def test_completely_different_title_not_matched(self):
        artist = self._make_artist("Band", ["Totally Different Album"])
        collection = ["After the Fire"]

        ranked = rank_artists_by_releases([artist], collection)

        assert ranked[0]["match_count"] == 0

    def test_each_discography_entry_used_at_most_once(self):
        artist = self._make_artist("Band", ["Same Title"])
        collection = ["Same Title", "Same Title"]

        ranked = rank_artists_by_releases([artist], collection)

        # Only one discography entry exists, so at most one can be matched
        assert ranked[0]["match_count"] == 1

    def test_empty_discography_yields_zero_matches(self):
        artist = {"name": "Band", "url": "https://ma.com", "discography": []}
        collection = ["After the Fire"]

        ranked = rank_artists_by_releases([artist], collection)

        assert ranked[0]["match_count"] == 0

    def test_empty_collection_yields_zero_matches(self):
        artist = self._make_artist("Band", ["After the Fire"])
        ranked = rank_artists_by_releases([artist], [])
        assert ranked[0]["match_count"] == 0

    def test_matched_release_entry_contains_expected_keys(self):
        artist = self._make_artist("Band", ["After the Fire"])
        ranked = rank_artists_by_releases([artist], ["After the Fire"])
        m = ranked[0]["matched_releases"][0]
        assert "collection_title" in m
        assert "discography_title" in m
        assert "discography_url" in m
        assert "discography_type" in m
        assert "discography_year" in m
        assert "match_score" in m

    def test_sorted_descending_by_match_count_then_avg_score(self):
        # a_high: 2 perfect matches; a_low: 1 fuzzy match
        a_high = self._make_artist("Band_A", ["After the Fire", "The Longest Night"])
        a_low = self._make_artist("Band_B", ["After the Fire"])
        collection = ["After the Fire", "The Longest Night"]

        ranked = rank_artists_by_releases([a_low, a_high], collection)

        assert ranked[0]["artist"]["name"] == "Band_A"
