# Tennis Schedule Generator

This project contains Python scripts to generate and manage a tennis schedule.

## Files

*   `tennis_schedule_maker_sanitized.py`: The main script to generate the schedule.
*   `sanitize_spreadsheet.py`: A utility script to sanitize the Excel file by removing personal information.
*   `tennis_sched.xlsx`: The Excel file containing the schedule data.
*   `schedule_output.txt`: The generated schedule.

## Setup

1.  Make sure you have Python 3 installed.
2.  Install the required library:
    ```bash
    pip install openpyxl
    ```

## Usage

### Generating the Schedule

1.  Update the `tennis_sched.xlsx` file with your schedule information.
2.  Run the script:
    ```bash
    python tennis_schedule_maker_sanitized.py
    ```
3.  The script will prompt you to enter the starting week number.
4.  The generated schedule will be saved in `schedule_output.txt`.