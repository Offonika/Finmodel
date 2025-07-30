from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from sheet_utils import apply_sheet_settings, hex_to_excel_tab_color

class FakeTab:
    def __init__(self):
        self.Color = None

class FakeApi:
    def __init__(self):
        self.Tab = FakeTab()
        self.moved = False
    def Move(self, Before=None, After=None):
        self.moved = True

class FakeSheet:
    def __init__(self, name, index=1):
        self.name = name
        self.index = index
        self.api = FakeApi()

class FakeSheets(list):
    def __getitem__(self, key):
        if isinstance(key, str):
            for s in self:
                if s.name == key:
                    return s
            raise KeyError(key)
        return list.__getitem__(self, key)

class FakeBook:
    def __init__(self, sheets):
        self.sheets = FakeSheets(sheets)


def test_apply_sheet_settings_color():
    sheet = FakeSheet('РасчетПлановыхПоказателей')
    wb = FakeBook([sheet])
    apply_sheet_settings(wb, 'РасчетПлановыхПоказателей')
    assert sheet.api.Tab.Color == hex_to_excel_tab_color('#D9D9D9')

