# Tennis Schedule Generat0r

This project contains a Python script to generate a tennis schedule from an Excel file.

## Files

*   `tennis_schedule_maker_sanitized.py`: The main script to generate the schedule.
*   `tennis_sched.xlsx`: The Excel file containing the schedule data.
*   `schedule_output.txt`: The generated schedule.

## Setup

1.  Make sure you have Python 3 installed.
2.  Install the required library:
    ```bash
    pip install openpyxl
    ```

## Configuration

Before running the script, you will need to update the following items in `tennis_schedule_maker_sanitized.py`:

1.  **Input File:** If your Excel file is not named `tennis_sched.xlsx`, change it in this line:
    ```python
    workbook = openpyxl.load_workbook('tennis_sched_sanitized.xlsx')
    ```

2.  **Output File:** If you want a different name for the output text file, change it in this line:
    ```python
    output_filename = 'schedule_output.txt'
    ```

3.  **Spreadsheet URL:** Replace the placeholder URL with your actual Google Sheets link in this line:
    ```python
    output_file.write("YOUR_SPREADSHEET_URL_HERE\n")
    ```

## Usage

1.  Update the `tennis_sched.xlsx` file with your schedule information.
2.  Run the script from your terminal:
    ```bash
    python tennis_schedule_maker_sanitized.py
    ```
3.  The script will prompt you to enter the starting week number.
4.  The generated schedule will be saved in `schedule_output.txt`.
