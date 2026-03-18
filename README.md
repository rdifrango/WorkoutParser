# WorkoutParser

A Python tool that parses monthly workout Excel files from [Chris Gates Fitness](https://www.chrisgatesfitness.com/) and consolidates all exercise data into a single output spreadsheet for easier tracking and analysis.

## Prerequisites

- Python 3.13+
- [uv](https://docs.astral.sh/uv/)

## Project Structure

```
WorkoutParser/
├── python/
│   ├── workout_parser/
│   │   ├── __init__.py
│   │   ├── __main__.py      # python -m support
│   │   ├── cli.py            # CLI entry point
│   │   └── parser.py         # Core parsing logic
│   ├── tests/
│   │   └── test_parser.py    # 14 tests
│   ├── pyproject.toml         # Project config & dependencies
│   └── uv.lock
└── workouts/                  # Input Excel files (not tracked in git)
```

## Setup

```bash
cd python
uv sync
```

## Usage

```bash
uv run python -m workout_parser.cli
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--folder`, `-f` | Path to folder containing workout `.xlsx` files | `workouts` |
| `--output`, `-o` | Output Excel file path | `output.xlsx` |

Example with a custom folder:

```bash
uv run python -m workout_parser.cli -f /path/to/files -o results.xlsx
```

## Testing

```bash
uv run pytest -v
```

## How It Works

1. Scans the input folder for `.xlsx` files
2. Parses the month and year from each filename (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx`)
3. Reads each "Week N" sheet, extracting daily exercises with sets, reps, and weight
4. Tracks per-day dates so exercises under each "Day N" header get the correct date
5. Writes all consolidated data to an output Excel file

### Expected Input Format

Each Excel file should follow the Chris Gates Fitness template:
- Sheets named "Week 1", "Week 2", etc.
- Daily exercises listed with "Day N" headers
- Exercise entries in `sets x reps x weight` format (e.g., `3x10x135`)

## Dependencies

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file reading
- [pandas](https://pandas.pydata.org/) — Data consolidation and Excel output