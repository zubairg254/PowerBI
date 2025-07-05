# FADL Load Processor GUI Tool (Revised Logic)

This Python-based GUI tool processes minute-wise load data from an Excel input file, performs calculations based on dispatch instructions and availability, and exports the results to a new Excel workbook. This version incorporates revised processing logic.

## Features

*   **Graphical User Interface (GUI):** Easy-to-use interface for file selection, configuration, and processing.
*   **Excel Input:** Reads data from a specified Excel file containing "Dispatch Instructions" and "Availability" sheets.
*   **Configurable Processing:**
    *   Select the specific month (mmm-yy format) to process from the input data (derived from "Dispatch Instructions" Column A).
    *   Choose a starting load type:
        *   **Custom Load:** Manually enter a starting load value (MW).
        *   **Use Final Availability (hourly from Col D):** Automatically fetches the starting load from Column D of the "Availability" sheet, based on the first hour of the selected month.
*   **Minute-by-Minute Calculation (Revised Logic):**
    *   Processes load data for every minute of the selected month.
    *   **Starting Load:** Determined by user selection (Custom or Final Availability from "Availability" Col D, first hour of month).
    *   **Dispatch Instruction Processing:**
        *   Instructions are read from the "Dispatch Instructions" sheet.
        *   **Ramp Calculation:**
            *   Ramp Rate (RR) = (`Target Demand (MW)` [Col F] - Previous Load) / `Ramp Duration (Minutes)` [Col C].
            *   `Ramp Duration (Minutes)` [Col C] is expected to be a positive number. If missing or invalid, it defaults to 1 minute with a warning.
            *   The load recorded for the minute of the instruction (`Notification Time` [Col A]) is the load *before* this new ramp begins. The first change in load due to the ramp is seen in the minute *following* the Notification Time.
            *   Load during ramp: `Previous Load + RR` per minute (effective from the minute after notification).
            *   Ramp continues until `Target Time Stamp` [Col B] is reached (preferred) or for the specified `Ramp Duration` [Col C] (if Col B is invalid/missing or not in the future relative to instruction time).
            *   Load is **not** capped at `Target Demand (MW)` during the ramp itself.
        *   **Post-Ramp / At Target Time Stamp:**
            *   If `Post-Ramp Target Type` [Col E] is "FCBL" (case-insensitive): The **`Load`** is set by looking up "Final Availability" (from "Availability" Col D) for the current hour. If this lookup fails (no data for the hour), the **`Load`** will be the `Target Demand (MW)` [Col F] that the ramp was aiming for (or was set to, if instantaneous).
            *   Otherwise (not "FCBL"): The **`Load`** is set to the `Target Demand (MW)` [Col F] that was ramped towards.
        *   **Instantaneous Changes:** An instruction results in an instantaneous load change if `Target Time Stamp` [Col B] is the same as `Notification Time` [Col A]. In this case, **`Load`** is set to `Target Demand (MW)` [Col F] at the `Notification Time`, and post-ramp logic (checking Col E for "FCBL") applies immediately for that same minute.
        *   **Between Instructions:** If not ramping and no new instruction, the previous minute's load is maintained.
*   **Excel Output:**
    *   Generates a new Excel workbook named `FADL Calculation.xlsx`.
    *   Output includes columns:
        *   `Date Time Stamp` (YYYY-MM-DD HH:MM:SS)
        *   `Load` (MW) - The calculated load for each minute.
        *   `Load Per Minute` (Load / 30)
        *   `Availability (hourly)` (MW): The hourly availability value (from Availability Sheet, Col D) corresponding to the hour of the `Date Time Stamp`. This value is constant for all minutes within the same hour. Displayed as blank/NA if no availability for that hour.
        *   `Demand (Load Demand MW)` (MW): The target demand set by the last active dispatch instruction (from Dispatch Instructions, Col F). This will be blank/NA if the system is currently following an "FCBL" post-ramp directive *and* the FCBL lookup was successful. If FCBL lookup fails, this column shows the fallback `Target Demand (MW)`.
        *   `LPM (30 Min Sum)`: This column is populated only at specific times:
            *   For rows where timestamp minute is `00` (e.g., `XX:00:00`): Value is the sum of 'Load Per Minute' from the previous 30 minutes (i.e., `(Hour-1):30:00` to `(Hour-1):59:00`).
            *   For rows where timestamp minute is `30` (e.g., `XX:30:00`): Value is the sum of 'Load Per Minute' from the first 30 minutes of the current hour (i.e., `Hour:00:00` to `Hour:29:00`).
            *   All other rows in this column will be blank/NA.
    *   Formatted output: Frozen headers, styled headers, auto-adjusted column widths, and appropriate number formatting.
*   **User Experience Features:**
    *   **Progress Bar:** Indicates the progress of the data processing.
    *   **Export Location Selection:** Allows the user to browse and select a folder to save the output file.
    *   **Status Log Window:** Provides real-time updates and messages about the tool's operations.

## Prerequisites

*   Python 3.x
*   The following Python libraries (also listed in `requirements.txt`):
    *   `pandas`
    *   `numpy` (for `pd.NA`)
    *   `openpyxl`
    *   `python-dateutil`

## Installation

1.  **Clone the repository or download the source code.**
2.  **Install Python 3.x** if you don't have it already.
3.  **Install required libraries:**
    Open a terminal or command prompt, navigate to the directory containing the tool's files (especially `requirements.txt`), and run:
    ```bash
    pip install -r requirements.txt
    ```

## How to Run

1.  Navigate to the directory where `load_processor_gui.py` is located using your terminal or command prompt.
2.  Run the script:
    ```bash
    python load_processor_gui.py
    ```
3.  The GUI application window will open.

## Input Excel File Format

The tool expects an Excel file (`.xlsx` or `.xls`) with the following structure:

1.  **Sheet Name: `Dispatch Instructions`**
    *   **Column A (Timestamp):** `Notification Time` - Date and Time Stamp of when the instruction becomes active. (Used for month selection).
    *   **Column B (Timestamp):** `Target Time Stamp` - The time by which the ramp should be completed and the target achieved. This is the primary determinant for ramp end if valid and in the future.
    *   **Column C (Numeric):** `Ramp Duration (Minutes)` - The duration in minutes over which the ramp should occur. Expected to be positive. Used if Column B is invalid/missing or not in the future. Defaults to 1 if missing/invalid and a ramp is implied.
    *   **Column D (Text/Any):** *Currently not used by the core processing logic.* Can be used for descriptive purposes by the user.
    *   **Column E (Text):** `Post-Ramp Target Type` - Specifies behavior after a ramp or instantaneous change. If this column contains the exact string "FCBL" (case-insensitive), the tool will look up Final Availability. Otherwise, the Target Demand is used.
    *   **Column F (Numeric):** `Target Demand (MW)` - The target load in MW that the instruction aims for.
    *   *Other columns can exist but are not currently used.*

2.  **Sheet Name: `Availability`**
    *   **Column A (Timestamp):** Date and Time Stamp.
    *   **Column D (Numeric):** `Final Availability (MW)` - The available load in MW. This column is used for:
        *   The "Use Final Availability (hourly from Col D)" starting load option (uses the first entry in the first hour of the selected month).
        *   The "FCBL" post-ramp logic (looks up the value for the current hour).
        *   Populating the 'Availability (hourly)' output column.
    *   *Columns B and C can exist but are not currently used by the core logic for these features.*

**Important Notes on Input Data:**
*   Ensure all timestamp columns are in a format pandas can recognize (e.g., `YYYY-MM-DD HH:MM:SS`).
*   `Ramp Duration (Minutes)` [Col C] is expected to be positive. If found to be missing or non-positive during parsing (and a ramp is necessary), it will default to 1 minute, and a warning will be logged.
*   The effect of a ramp (change in load due to ramp rate) starts from the minute *following* the `Notification Time`. The load recorded *at* the `Notification Time` is the load just before the new ramp is initiated.
*   The "FCBL" string in Column E of "Dispatch Instructions" must be exact (though it's processed case-insensitively) to trigger the hourly availability lookup. If "FCBL" is specified and lookup fails, the load defaults to the `Target Demand (MW)` from Column F.
*   Hourly lookups in the "Availability" sheet (Column D) use the *first valid numeric entry* found within the specified hour.

## Using the Tool

1.  **Select Excel File:** Click "Browse..." next to "Excel File:", choose your input Excel file, and click "Open".
    *   The tool will read the file. If successful, the "Month to Process" dropdown will be populated. "Start Processing" button enabled if data seems valid.
    *   Errors (missing sheets, critical data format issues) will be shown in messages and the Status Log.
2.  **Select Month to Process:** Choose the desired month from the dropdown.
3.  **Select Starting Load Type:**
    *   **Use Final Availability (hourly from Col D):** (Default) The tool attempts to get the starting load from Column D of the "Availability" sheet (first entry in the first hour of the selected month).
    *   **Custom Load:** Select this and enter a numeric starting load in the entry field.
4.  **Select Export Folder:** Click "Browse..." to choose the save location for `FADL Calculation.xlsx`.
5.  **Start Processing:** Click "Start Processing".
    *   The button is disabled during processing. Progress and status are updated.
6.  **Output:** On completion, a success message appears. `FADL Calculation.xlsx` is saved. Errors are reported.

## Troubleshooting

*   **"Sheet not found" / "Insufficient columns":** Verify sheet names ("Dispatch Instructions", "Availability") and that they have the minimum required columns as specified above.
*   **Date/Numeric Conversion Errors:** Check data types in relevant columns. Timestamps should be clear date/time formats. Numeric columns (Ramp Duration, Target Demand, Final Availability) must contain numbers. The Status Log often provides details on conversion issues.
*   **Initial Hourly Availability / FCBL Lookup Failures:**
    *   Ensure the "Availability" sheet has data in Column D for the relevant hour(s).
    *   Verify Column A (timestamps) and Column D (load values) in "Availability" are correctly formatted.
*   **PermissionError on save:** Ensure write permissions for the export folder and that the output file isn't open.

---
This README reflects the new logic and input requirements.
