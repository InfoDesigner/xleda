from pathlib import Path

import pandas as pd
import pytest
import sys

from xleda import os_interface
from xleda.utilities import DataSetParser


# TODO: Change these 
def test_read_dataframe_from_csv(tmp_path: Path):
    csv_path = tmp_path / "sample.csv"
    expected = pd.DataFrame({"a": [1, 2], "b": ["x", "y"]})
    expected.to_csv(csv_path, index=False)

    actual = DataSetParser(csv_path)

    pd.testing.assert_frame_equal(actual, expected)


def test_create_workbook_from_file_uses_file_stem_and_parent(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    csv_path = tmp_path / "sales data.csv"
    pd.DataFrame({"amount": [10, 20]}).to_csv(csv_path, index=False)
    captured = {}

    def fake_wb(**kwargs):
        captured.update(kwargs)
        return "created"

    monkeypatch.setattr(os_interface, "wb", fake_wb)

    result = os_interface.create_workbook_from_file(csv_path, overwrite=True, no_vba=True)

    assert result == "created"
    assert captured["name"] == "sales data"
    assert captured["wb_path"] == tmp_path
    assert captured["overwrite"] is True
    assert captured["no_vba"] is True
    pd.testing.assert_frame_equal(captured["input_df"], pd.DataFrame({"amount": [10, 20]}))


def test_unsupported_file_type_raises(tmp_path: Path):
    unsupported_path = tmp_path / "sample.json"
    unsupported_path.write_text("{}", encoding="utf-8")

    with pytest.raises(ValueError, match="Unsupported file type"):
        os_interface.create_workbook_from_file(unsupported_path)

@pytest.mark.skipif(sys.platform == 'win32', reason="Does not work on Windows")
def test_macos_workflow_opens_terminal():
    script = os_interface.macos_workflow_shell_script()

    assert "Terminal" in script
    assert "-m xleda wb" in script


@pytest.mark.skipif(sys.platform == 'darwin', reason="Does not work on macOS")
def test_windows_command_opens_powershell():
    command = os_interface.windows_command()

    assert "powershell.exe" in command
    assert "-m xleda wb" in command
    assert "%1" in command
