# WorkoutParser

A Java tool that parses monthly workout Excel files from [Chris Gates Fitness](https://www.chrisgatesfitness.com/) and consolidates all exercise data into a single output spreadsheet for easier tracking and analysis.

## Prerequisites

- Java 21+
- Maven 3.x

## Project Structure

```
WorkoutParser/
├── src/main/java/org/difrango/
│   └── WorkoutParser.java       # Main application
├── workouts/                    # Input Excel files (not tracked in git)
├── pom.xml                      # Maven build configuration
└── .mvn/jvm.config              # JVM args for Error Prone
```

## Build

```bash
mvn clean compile
```

## Usage

```bash
mvn exec:java -Dexec.mainClass="org.difrango.WorkoutParser"
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--folder`, `-f` | Path to folder containing workout `.xlsx` files | `workouts` |

Example with a custom folder:

```bash
mvn exec:java -Dexec.mainClass="org.difrango.WorkoutParser" -Dexec.args="--folder /path/to/files"
```

## How It Works

1. Scans the input folder for `.xlsx` files
2. Parses the month and year from each filename (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx`)
3. Reads each "Week N" sheet, extracting daily exercises with sets, reps, and weight
4. Writes all consolidated data to `output.xlsx`

### Expected Input Format

Each Excel file should follow the Chris Gates Fitness template:
- Sheets named "Week 1", "Week 2", etc.
- Daily exercises listed with "Day N" headers
- Exercise entries in `sets x reps x weight` format (e.g., `3x10x135`)

## Dependencies

- [Apache POI](https://poi.apache.org/) — Excel file reading/writing
- [JCommander](https://jcommander.org/) — CLI argument parsing
- [Error Prone](https://errorprone.info/) — Static analysis at compile time