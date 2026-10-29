from main import load_settings, save_settings


def test_settings_round_trip(tmp_path):
    path = tmp_path / "SynapseSorter" / "settings.json"
    save_settings({"header_path": "C:/header.docx", "header_values": {"Midterm": "Final"}}, str(path))
    assert load_settings(str(path)) == {"header_path": "C:/header.docx", "header_values": {"Midterm": "Final"}}


def test_missing_or_broken_settings_start_empty(tmp_path):
    assert load_settings(str(tmp_path / "missing.json")) == {}
    broken = tmp_path / "broken.json"
    broken.write_text("{not json")
    assert load_settings(str(broken)) == {}
