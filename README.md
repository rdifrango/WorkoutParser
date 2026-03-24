# WorkoutParser

A Python tool that parses monthly workout spreadsheet files (Excel and Apple Numbers) from [Chris Gates Fitness](https://www.chrisgatesfitness.com/) and consolidates all exercise data into a single output spreadsheet for easier tracking and analysis.

## Prerequisites

- Python 3.13+
- [uv](https://docs.astral.sh/uv/)

## Project Structure

```
WorkoutParser/
├── workout_parser/
│   ├── __init__.py
│   ├── __main__.py          # python -m support
│   ├── cli.py                # CLI entry point
│   └── parser.py             # Core parsing logic
├── tests/
│   └── test_parser.py        # 14 tests
├── streamlit_app.py           # Streamlit web app
├── pyproject.toml             # Project config & dependencies
├── requirements.txt           # For Streamlit Community Cloud
├── uv.lock
└── workouts/                  # Input Excel files (not tracked in git)
```

## Setup

```bash
uv sync
```

## Usage

```bash
uv run python -m workout_parser.cli
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--folder`, `-f` | Path to folder containing workout `.xlsx` / `.numbers` files | `workouts` |
| `--output`, `-o` | Output Excel file path | `output.xlsx` |

Example with a custom folder:

```bash
uv run python -m workout_parser.cli -f /path/to/files -o results.xlsx
```

### Web App

The app is live at **[gates-workout-parser.streamlit.app](https://gates-workout-parser.streamlit.app/)**.

To run locally:

```bash
uv sync
uv run streamlit run streamlit_app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

## Testing

```bash
uv run pytest -v
```

## How It Works

1. Scans the input folder for `.xlsx` and `.numbers` files
2. Parses the month and year from each filename (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx` or `.numbers`)
3. Reads each "Week N" sheet, extracting daily exercises with sets, reps, and weight
4. Tracks per-day dates so exercises under each "Day N" header get the correct date
5. Writes all consolidated data to an output Excel file

### Expected Input Format

Each spreadsheet file should follow the Chris Gates Fitness template:
- Filename in the format `Month-Year-...` (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx` or `.numbers`)
- Sheets named "Week 1", "Week 2", etc.
- Daily exercises listed with "Day N" headers
- Exercise data is read from the **Client Notes** section
- Sets/reps/weight in the format `{sets}x{reps}x{weight}` (e.g., `3x10x135`)

## Dependencies

- [numbers-parser](https://github.com/masaccio/numbers-parser) — Apple Numbers file reading
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file reading
- [pandas](https://pandas.pydata.org/) — Data consolidation and Excel output
- [streamlit](https://streamlit.io/) — Web app frontend