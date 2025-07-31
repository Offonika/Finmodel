from __future__ import annotations

import os
from pathlib import Path


def _parse_xlwings_conf(path: Path) -> str | None:
    """Return interpreter path from ``.xlwings.conf`` if present."""
    conf: dict[str, str] = {}
    for line in path.read_text().splitlines():
        if "=" not in line:
            continue
        key, val = line.split("=", 1)
        conf[key.strip()] = val.split(";", 1)[0].strip()
    interp = conf.get("INTERPRETER")
    if interp and "%(PROJECT_PATH)s" in interp:
        interp = interp.replace("%(PROJECT_PATH)s", conf.get("PROJECT_PATH", ""))
    return interp


def ensure_interpreter_path() -> None:
    """Raise ``FileNotFoundError`` if ``INTERPRETER`` path is invalid on Windows."""
    if os.name != "nt":
        return

    conf_path = Path(__file__).resolve().parents[1] / ".xlwings.conf"
    if not conf_path.exists():
        return

    interpreter = _parse_xlwings_conf(conf_path)
    if interpreter and not Path(interpreter).exists():
        raise FileNotFoundError(
            f"Interpreter not found: {interpreter}. "
            "Run setup again or update .xlwings.conf."
        )
