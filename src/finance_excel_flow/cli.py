from __future__ import annotations

import argparse
import json
from pathlib import Path

from .config import load_config
from .engine import describe_config, process_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="finance-excel-flow",
        description="Process monthly Excel files using a YAML configuration.",
    )
    parser.add_argument("--input", required=True, help="Path to the source Excel file")
    parser.add_argument("--config", required=True, help="Path to the YAML configuration")
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Directory for the generated file. Defaults to ./output next to the source file.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Optional explicit output file path. Overrides --output-dir.",
    )
    parser.add_argument(
        "--print-config",
        action="store_true",
        help="Print the normalized configuration as JSON and exit.",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    config = load_config(args.config)
    if args.print_config:
        print(json.dumps(describe_config(config), ensure_ascii=False, indent=2, default=str))
        return

    result_path = process_file(
        input_path=Path(args.input),
        config=config,
        output_dir=Path(args.output_dir) if args.output_dir else None,
        output_path=Path(args.output) if args.output else None,
    )
    print(str(result_path))

