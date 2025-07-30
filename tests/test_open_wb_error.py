import sys
from pathlib import Path
import pytest

sys.path.append(str(Path(__file__).resolve().parents[1] / "scripts"))
import fill_planned_indicators


def test_open_wb_missing(tmp_path, monkeypatch):
    fake_path = tmp_path / "missing.xlsm"
    monkeypatch.setattr(fill_planned_indicators, "EXCEL_PATH", fake_path)
    with pytest.raises(FileNotFoundError):
        fill_planned_indicators.open_wb()
