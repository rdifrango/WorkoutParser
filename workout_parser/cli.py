"""Command-line interface for workout-parser."""

import argparse
from pathlib import Path

from workout_parser.parser import parse_folder, write_output


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Parse workout Excel files and consolidate into a single spreadsheet."
    )
    parser.add_argument(
        "-f",
        "--folder",
        type=Path,
        default=Path("workouts"),
        help="Path to folder containing workout .xlsx files (default: workouts)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("output.xlsx"),
        help="Output Excel file path (default: output.xlsx)",
    )
    args = parser.parse_args()

    if not args.folder.is_dir():
        parser.error(f"Folder not found: {args.folder}")

    df = parse_folder(args.folder)
    print(f"Parsed {len(df)} exercises from {args.folder}")
    write_output(df, args.output)
    print(f"Written to {args.output}")


if __name__ == "__main__":
    main()