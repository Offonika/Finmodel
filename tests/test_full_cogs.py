from pathlib import Path
import sys
sys.path.append(str(Path(__file__).resolve().parents[1] / 'scripts'))
from fill_planned_indicators import full_cogs

def test_full_cogs_reduced_rates():
    assert full_cogs(100, 5) == 105
    assert full_cogs(200, 7) == 214
