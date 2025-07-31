# main.py

from __future__ import annotations

import argparse
import configparser
from pathlib import Path

from file_loader import load_files
from aggregator import aggregate_data
from excel_writer import write_to_excel, write_df_to_excel_table

BASE_DIR = Path(__file__).resolve().parents[1]
CONFIG_PATH = BASE_DIR / "config.ini"


def read_config() -> tuple[str | None, str | None]:
    """Read default paths from ``config.ini`` if present."""
    if not CONFIG_PATH.exists():
        return None, None

    parser = configparser.ConfigParser()
    parser.read(CONFIG_PATH)
    org_folder = parser.get("paths", "org_folder", fallback=None)
    output_path = parser.get("paths", "output_path", fallback=None)
    return org_folder, output_path

def main(args: list[str] | None = None) -> None:
    """Aggregate files from ``org_folder`` and write into ``output_path``."""

    cfg_org, cfg_out = read_config()

    parser = argparse.ArgumentParser(description="Aggregate Ozon service charges")
    parser.add_argument(
        "--org_folder",
        type=str,
        default=cfg_org,
        help="Folder with source organization files",
    )
    parser.add_argument(
        "--output_path",
        type=str,
        default=cfg_out,
        help="Path to output Finmodel workbook",
    )

    parsed = parser.parse_args(args)

    if parsed.org_folder is None or parsed.output_path is None:
        parser.error("Both org_folder and output_path must be provided")

    org_folder = (BASE_DIR / Path(parsed.org_folder)).expanduser()
    output_path = (BASE_DIR / Path(parsed.output_path)).expanduser()

    files_df = load_files(str(org_folder))
    result_df = aggregate_data(files_df)

    sheet = "НачисленияУслугОзон"
    table = "НачисленияУслугОзонTable"

    # 1. Сохраняем просто DataFrame на лист
    write_to_excel(result_df, str(output_path), sheet_name=sheet)
    # 2. Преобразуем лист в умную таблицу (Excel Table)
    write_df_to_excel_table(result_df, str(output_path), sheet, table)

if __name__ == '__main__':
    main()
