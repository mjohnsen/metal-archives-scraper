import json
from unittest.mock import patch

import pytest

import metal_archives_scraper.config as config_module
from metal_archives_scraper.config import load_config, save_config, get_spreadsheet_path


@pytest.fixture(autouse=True)
def isolated_config(tmp_path, monkeypatch):
    """Redirect CONFIG_FILE to a temp path so tests never touch the real config."""
    monkeypatch.setattr(config_module, "CONFIG_FILE", tmp_path / "test_config.json")


class TestLoadConfig:
    def test_returns_empty_dict_when_file_absent(self):
        assert load_config() == {}

    def test_returns_parsed_json_when_file_exists(self, tmp_path, monkeypatch):
        data = {"spreadsheet_path": "/some/file.xlsx", "extra": 42}
        cfg_file = tmp_path / "cfg.json"
        cfg_file.write_text(json.dumps(data))
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)

        assert load_config() == data


class TestSaveConfig:
    def test_writes_valid_json(self, tmp_path, monkeypatch):
        cfg_file = tmp_path / "cfg.json"
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)
        save_config({"key": "value"})

        assert json.loads(cfg_file.read_text()) == {"key": "value"}

    def test_overwrites_existing_file(self, tmp_path, monkeypatch):
        cfg_file = tmp_path / "cfg.json"
        cfg_file.write_text(json.dumps({"old": "data"}))
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)
        save_config({"new": "data"})

        assert json.loads(cfg_file.read_text()) == {"new": "data"}


class TestGetSpreadsheetPath:
    def test_returns_saved_path_when_file_exists(self, tmp_path, monkeypatch):
        xlsx = tmp_path / "collection.xlsx"
        xlsx.touch()
        cfg_file = tmp_path / "cfg.json"
        cfg_file.write_text(json.dumps({"spreadsheet_path": str(xlsx)}))
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)

        assert get_spreadsheet_path() == str(xlsx)

    def test_prompts_when_saved_path_is_missing(self, tmp_path, monkeypatch, capsys):
        real_xlsx = tmp_path / "real.xlsx"
        real_xlsx.touch()
        cfg_file = tmp_path / "cfg.json"
        cfg_file.write_text(json.dumps({"spreadsheet_path": "/does/not/exist.xlsx"}))
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)

        with patch("builtins.input", return_value=str(real_xlsx)):
            result = get_spreadsheet_path()

        assert result == str(real_xlsx)
        assert "not found" in capsys.readouterr().out

    def test_prompts_when_no_config_file(self, tmp_path, monkeypatch):
        xlsx = tmp_path / "collection.xlsx"
        xlsx.touch()
        with patch("builtins.input", return_value=str(xlsx)):
            result = get_spreadsheet_path()

        assert result == str(xlsx)

    def test_strips_surrounding_quotes_from_input(self, tmp_path, monkeypatch):
        xlsx = tmp_path / "collection.xlsx"
        xlsx.touch()
        with patch("builtins.input", return_value=f'"{xlsx}"'):
            result = get_spreadsheet_path()

        assert result == str(xlsx)

    def test_retries_until_valid_path_given(self, tmp_path, monkeypatch):
        real_xlsx = tmp_path / "collection.xlsx"
        real_xlsx.touch()
        inputs = iter(["/nonexistent/path.xlsx", str(real_xlsx)])
        with patch("builtins.input", side_effect=inputs):
            result = get_spreadsheet_path()

        assert result == str(real_xlsx)

    def test_saves_valid_path_to_config(self, tmp_path, monkeypatch):
        xlsx = tmp_path / "collection.xlsx"
        xlsx.touch()
        cfg_file = tmp_path / "cfg.json"
        monkeypatch.setattr(config_module, "CONFIG_FILE", cfg_file)

        with patch("builtins.input", return_value=str(xlsx)):
            get_spreadsheet_path()

        saved = json.loads(cfg_file.read_text())
        assert saved["spreadsheet_path"] == str(xlsx)
