import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from law_config import load_oc, save_oc


def test_load_oc_returns_none_when_file_missing(tmp_path):
    path = tmp_path / "config_법률검토.json"
    assert load_oc(str(path)) is None


def test_save_and_load_oc_roundtrip(tmp_path):
    path = tmp_path / "config_법률검토.json"
    save_oc("hong", str(path))
    assert load_oc(str(path)) == "hong"


def test_load_oc_returns_none_on_corrupt_json(tmp_path):
    path = tmp_path / "config_법률검토.json"
    path.write_text("not json", encoding="utf-8")
    assert load_oc(str(path)) is None
