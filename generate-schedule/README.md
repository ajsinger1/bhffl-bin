# BHFFL Fantasy Football Schedule Generator

## Overview

The BHFFL Fantasy Football Schedule Generator creates a 14-week schedule for a fantasy football league, ensuring all teams play according to set rules. The script uses a backtracking algorithm with MRV and forward-checking heuristics and can output the schedule in CSV or XLSX format.

## Usage

Run the script with:

```bash
python generate_schedule.py -f path/to/divisions.csv -o output_schedule.xlsx
```

### Options

- **`-f, --file`**: Path to the input CSV file (default: `example_divisions.csv`).
- **`-o, --output`**: Path to save the schedule (csv/xlsx) (default: `schedule.xlsx`).
- **`-n, --no-output`**: Skip saving the output file.
- **`-i, --info`**: Display script information.

### Example

Generate a schedule using the example CSV file:

```bash
python generate_schedule.py -f example_divisions.csv -o schedule.xlsx
```

## Input Format

The CSV file needs one column per division (`DIVISION 1`, `DIVISION 2`, `DIVISION 3`), each
holding four team names — 12 teams in total. See `example_divisions.csv`.

## Schedule Details

- **Divisions**: Three divisions of four teams.
- **In-Division Games**: Teams play each division rival twice (6 games).
- **Out-of-Division Games**: Teams play all other teams once (8 games).
- **Season Length**: 14 weeks, six games per week.
- **Randomized**: The schedule is shuffled each time.

## License

This project is licensed under the MIT License.
