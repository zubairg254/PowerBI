# FADL Load Processor GUI Tool

This Python-based GUI tool processes minute-wise load data from an Excel input file, performs calculations based on dispatch instructions and availability, and exports the results to a new Excel workbook.

## Features

*   **Graphical User Interface (GUI):** Easy-to-use interface for file selection, configuration, and processing.
*   **Excel Input:** Reads data from a specified Excel file containing "Dispatch Instructions" and "Availability" sheets.
*   **Configurable Processing:**
    *   Select the specific month (mmm-yy format) to process from the input data.
    *   Choose a starting load type:
        *   **Custom Load:** Manually enter a starting load value (MW).
        *   **FCBL (First Cleared Bid Load):** Automatically fetches the load from the first hour of availability data for the selected month.
*   **Minute-by-Minute Calculation:**
    *   Processes load data for every minute of the selected month.
    *   Applies logic for:
        *   Initial load (Custom or FCBL).
        *   Ramp rates as per dispatch instructions (ramping is not capped by target loads).
        *   Maintaining previous load or using "Final Availability" data when between instructions.
*   **Excel Output:**
    *   Generates a new Excel workbook named `FADL Calculation.xlsx`.
    *   Output includes:
        *   `Date Time Stamp` (YYYY-MM-DD HH:MM:SS)
        *   `Load` (MW)
        *   `Load Per Minute` (Load / 30)
        *   `LPM (30 Min Sum)` (rolling sum of "Load Per Minute" for each 30-minute window)
    *   Formatted output: Frozen headers, styled headers, auto-adjusted column widths, and appropriate number formatting.
*   **User Experience Features:**
    *   **Progress Bar:** Indicates the progress of the data processing.
    *   **Export Location Selection:** Allows the user to browse and select a folder to save the output file.
    *   **Status Log Window:** Provides real-time updates and messages about the tool's operations (e.g., file reading, processing status, errors, save location).

## Prerequisites

*   Python 3.x
*   The following Python libraries (also listed in `requirements.txt`):
    *   `pandas`
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
    *   **Column A:** Date and Time Stamp of the instruction (must be convertible to datetime objects, e.g., `YYYY-MM-DD HH:MM:SS` or similar). This column is used to determine available months for processing.
    *   **Column B (Assumed):** Notification Type (Text, e.g., `RAMP`, `TARGET`). Case-insensitive, leading/trailing spaces are ignored.
    *   **Column C (Assumed):** Value (Numeric).
        *   If Notification Type is `RAMP`, this is the ramp rate in MW/minute.
        *   If Notification Type is `TARGET`, this is the target load in MW.
    *   *Other columns can exist but are not currently used by the core logic.*

2.  **Sheet Name: `Availability`**
    *   **Column A (Assumed):** Date and Time Stamp (must be convertible to datetime objects).
    *   **Column B (Assumed):** Available Load (Numeric, in MW). This data is used for:
        *   Determining FCBL (first hour's availability of the selected month).
        *   Potentially for the "Final Availability" logic when between dispatch instructions (currently, the tool primarily maintains previous load if not ramping and no exact match is found in this sheet for the current minute).
    *   *Other columns can exist but are not currently used.*

**Important Notes on Input Data:**
*   Ensure timestamps in both sheets are consistent and cover the desired processing periods.
*   The interpretation of "Notification Type" and "Value" columns in "Dispatch Instructions" is based on the assumed string values (`RAMP`, `TARGET`). Other types will be logged as "Unknown".
*   The "Final Availability" logic currently looks for an exact timestamp match in the Availability sheet to determine load when between instructions and not ramping. If no match is found, the previous load is maintained.

## Using the Tool

1.  **Select Excel File:** Click "Browse..." next to "Excel File:", choose your input Excel file, and click "Open".
    *   The tool will attempt to read the file. If successful, the "Month to Process" dropdown will be populated.
    *   If there are errors (missing sheets, bad data format), an error message will appear, and the "Start Processing" button may be disabled. Check the Status Log for details.
2.  **Select Month to Process:** Choose the desired month from the dropdown list.
3.  **Select Starting Load Type:**
    *   **FCBL:** The tool will try to get the starting load from the "Availability" sheet for the first hour of the selected month.
    *   **Custom Load:** Select this option and enter a numeric value for the starting load in the adjacent entry field.
4.  **Select Export Folder:** Click "Browse..." next to "Export Folder:", choose where you want to save the `FADL Calculation.xlsx` output file, and click "Select Folder".
5.  **Start Processing:** Click the "Start Processing" button.
    *   The button will be disabled during processing.
    *   The Progress Bar will show the progress.
    *   The Status Log will display real-time updates.
6.  **Output:** Once processing is complete, a success message will appear, and the `FADL Calculation.xlsx` file will be saved in the chosen export folder. If errors occur, they will be reported in a messagebox and/or the Status Log.

## Troubleshooting

*   **"Sheet not found" errors:** Ensure your Excel file contains sheets named exactly "Dispatch Instructions" and "Availability".
*   **Date conversion errors:** Check that the first column in both sheets contains valid date/time information that pandas can recognize.
*   **FCBL errors:**
    *   Ensure the "Availability" sheet has data, especially for the first hour of the month you are processing.
    *   Verify the first column is timestamps and the second column contains numeric load values.
*   **PermissionError on save:** Make sure the application has write permissions to the selected export folder and that `FADL Calculation.xlsx` is not already open in Excel.
*   **GUI Freezes (should not happen):** If the GUI freezes, it might indicate an unexpected issue with the threading. Please report this. The Status Log (if still updating or via console output) might provide clues.

---

This README provides a good overview for users.
