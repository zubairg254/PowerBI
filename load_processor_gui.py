import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import threading
import datetime
from dateutil.relativedelta import relativedelta

class LoadProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FADL Load Processor")
        self.root.geometry("800x700") # Increased height for log
        self.root.minsize(600, 500)

        # --- Variables ---
        self.file_path_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.start_load_type_var = tk.StringVar(value="FCBL") # Default to FCBL
        self.custom_load_var = tk.DoubleVar(value=0.0)
        self.export_dir_var = tk.StringVar()
        self.df_dispatch = None
        self.df_availability = None

        # --- Main Frame ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # --- File Selection ---
        file_frame = ttk.LabelFrame(main_frame, text="Input File", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).grid(row=0, column=2, sticky=tk.E, padx=5, pady=5)

        # --- Configuration ---
        config_frame = ttk.LabelFrame(main_frame, text="Processing Configuration", padding="10")
        config_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Month to Process:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.month_combo = ttk.Combobox(config_frame, textvariable=self.month_var, state="readonly", width=15)
        self.month_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.month_combo['values'] = ["Select a file first"]

        ttk.Label(config_frame, text="Starting Load Type:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        fcbl_radio = ttk.Radiobutton(config_frame, text="FCBL", variable=self.start_load_type_var, value="FCBL", command=self.toggle_custom_load_entry)
        fcbl_radio.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        custom_load_frame = ttk.Frame(config_frame)
        custom_load_frame.grid(row=2, column=1, sticky=(tk.W, tk.E))

        custom_radio = ttk.Radiobutton(custom_load_frame, text="Custom Load:", variable=self.start_load_type_var, value="Custom", command=self.toggle_custom_load_entry)
        custom_radio.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.custom_load_entry = ttk.Entry(custom_load_frame, textvariable=self.custom_load_var, width=10, state=tk.DISABLED)
        self.custom_load_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        # --- Export Location ---
        export_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        export_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        export_frame.columnconfigure(1, weight=1)

        ttk.Label(export_frame, text="Export Folder:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(export_frame, textvariable=self.export_dir_var, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(export_frame, text="Browse...", command=self.browse_export_dir).grid(row=0, column=2, sticky=tk.E, padx=5, pady=5)

        # --- Controls & Progress ---
        run_frame = ttk.Frame(main_frame, padding="10")
        run_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        run_frame.columnconfigure(0, weight=1) # Make progress bar expand

        self.process_button = ttk.Button(run_frame, text="Start Processing", command=self.start_processing_thread, width=20)
        self.process_button.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(run_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

        # --- Status Log ---
        log_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1) # Make text area expand

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text['yscrollcommand'] = log_scroll.set
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # --- Initial Status ---
        self.status_log("Application started. Please select an Excel file.")
        self.toggle_custom_load_entry() # Set initial state of custom load entry

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_log(f"Selected input file: {file_path}")
            # Attempt to read data and populate months
            # For now, this uses the placeholder read_excel_data
            # In the future, this will store the actual df_dispatch and df_availability
            _, _, available_months = self.read_excel_data(file_path)
            if available_months:
                self.month_combo['values'] = available_months
                if available_months:
                    self.month_var.set(available_months[0]) # Select first month by default
                else:
                    self.month_var.set("")
                self.status_log(f"Available months populated: {available_months}")
            else:
                self.month_combo['values'] = ["Error reading file"]
                self.month_var.set("")
                self.status_log("Could not populate months. Error reading file or no valid months found.")
                self.process_button.config(state=tk.DISABLED) # Disable processing

            # Also check if dataframes themselves were loaded successfully
            if self.df_dispatch is None or self.df_availability is None:
                self.status_log("One or both required data sheets were not loaded successfully. Processing disabled.")
                self.process_button.config(state=tk.DISABLED)
                if not available_months: # If months were already empty, don't overwrite combobox again
                    self.month_combo['values'] = ["Error in sheet data"]
                    self.month_var.set("")
            elif available_months: # Only enable if DFs are good AND months are available
                 self.process_button.config(state=tk.NORMAL)

        else:
            self.status_log("File selection cancelled.")
            # Consider if process button should be disabled if no file is selected at all
            # self.process_button.config(state=tk.DISABLED)


    def browse_export_dir(self):
        dir_path = filedialog.askdirectory(title="Select Export Directory")
        if dir_path:
            self.export_dir_var.set(dir_path)
            self.status_log(f"Selected export directory: {dir_path}")
        else:
            self.status_log("Export directory selection cancelled.")

    def toggle_custom_load_entry(self):
        if self.start_load_type_var.get() == "Custom":
            self.custom_load_entry.config(state=tk.NORMAL)
        else:
            self.custom_load_entry.config(state=tk.DISABLED)
            self.custom_load_var.set(0.0) # Reset when disabled
        self.status_log(f"Start load type set to: {self.start_load_type_var.get()}")

    # --- Placeholder Core Logic Functions ---
    def read_excel_data(self, file_path):
        """Reads data from the Excel file and extracts available months."""
        self.status_log(f"Reading Excel file: {file_path}")
        # Reset dataframes
        self.df_dispatch = None
        self.df_availability = None
        available_months = []

        if not file_path:
            # This case should ideally be caught before calling, but good to have.
            messagebox.showerror("Error", "No file selected to read.")
            self.status_log("File reading skipped: No file path provided.")
            return self.df_dispatch, self.df_availability, []

        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names

            required_sheets = {"Dispatch Instructions", "Availability"}
            if not required_sheets.issubset(sheet_names):
                missing_sheets = required_sheets - set(sheet_names)
                err_msg = f"Missing required sheet(s): {', '.join(missing_sheets)}."
                messagebox.showerror("Sheet Error", err_msg + f"\nFound sheets: {sheet_names}")
                self.status_log(f"Error: {err_msg} Excel file only contains: {sheet_names}")
                return self.df_dispatch, self.df_availability, []

            self.status_log(f"Found sheets: {sheet_names}. Reading 'Dispatch Instructions' and 'Availability'.")

            # Read Dispatch Instructions
            self.df_dispatch = pd.read_excel(xls, sheet_name="Dispatch Instructions")
            self.status_log(f"'Dispatch Instructions' sheet read. Rows: {len(self.df_dispatch)}, Columns: {len(self.df_dispatch.columns)}")

            if self.df_dispatch.empty:
                self.status_log("Warning: 'Dispatch Instructions' sheet is empty. No months to process.")
            elif self.df_dispatch.columns.empty:
                 self.status_log("Warning: 'Dispatch Instructions' sheet has no columns.")
                 messagebox.showwarning("Data Warning", "'Dispatch Instructions' sheet has no columns.")
                 self.df_dispatch = None # Invalidate
            elif not pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                self.status_log("First column of 'Dispatch Instructions' is not datetime. Attempting conversion...")
                try:
                    self.df_dispatch.iloc[:, 0] = pd.to_datetime(self.df_dispatch.iloc[:, 0], errors='coerce')
                    if self.df_dispatch.iloc[:, 0].isnull().all() and not self.df_dispatch.empty: # All values failed to convert
                         messagebox.showerror("Data Error", "First column of 'Dispatch Instructions' (expected dates) could not be converted to dates/timestamps. All values are invalid.")
                         self.status_log("Error: All values in first column of 'Dispatch Instructions' failed datetime conversion.")
                         self.df_dispatch = None # Invalidate
                    elif self.df_dispatch.iloc[:, 0].isnull().any():
                        num_failed = self.df_dispatch.iloc[:,0].isnull().sum()
                        self.status_log(f"Warning: {num_failed} values in the first column of 'Dispatch Instructions' failed datetime conversion and were set to NaT.")
                        messagebox.showwarning("Data Conversion Warning", f"{num_failed} date entries in 'Dispatch Instructions' are invalid and were ignored.")
                    else:
                         self.status_log("Successfully converted first column of 'Dispatch Instructions' to datetime.")
                except Exception as e_conv:
                    messagebox.showerror("Data Error", f"Could not convert first column of 'Dispatch Instructions' to dates/timestamps.\nError: {e_conv}")
                    self.status_log(f"Error converting first column of 'Dispatch Instructions': {e_conv}")
                    self.df_dispatch = None # Invalidate

            if self.df_dispatch is not None and not self.df_dispatch.empty and pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                # Get months from non-NaT dates only
                valid_dates = self.df_dispatch.iloc[:, 0].dropna()
                if not valid_dates.empty:
                    available_months = sorted(valid_dates.dt.strftime('%b-%y').unique().tolist())

                if not available_months:
                    self.status_log("No valid months found in 'Dispatch Instructions' (Column A after processing).")
                    messagebox.showwarning("Data Warning", "No processable months found in 'Dispatch Instructions' (Column A). Ensure it contains valid dates in the first column.")

            # Read Availability
            self.df_availability = pd.read_excel(xls, sheet_name="Availability")
            self.status_log(f"'Availability' sheet read. Rows: {len(self.df_availability)}, Columns: {len(self.df_availability.columns)}")

            if self.df_availability.empty:
                 self.status_log("Warning: 'Availability' sheet is empty. FCBL option might not work if selected.")
            elif self.df_availability.columns.empty:
                self.status_log("Warning: 'Availability' sheet has no columns.")
                messagebox.showwarning("Data Warning", "'Availability' sheet has no columns.")
                self.df_availability = None # Invalidate

            # Return success flags based on whether dataframes were loaded.
            # The actual dataframes are stored in self.df_dispatch and self.df_availability.
            return self.df_dispatch is not None, self.df_availability is not None, available_months

        except FileNotFoundError:
            messagebox.showerror("File Error", f"File not found: {file_path}")
            self.status_log(f"Error: File not found at {file_path}")
            return None, None, [] # Using None to indicate critical failure
        except ValueError as ve: # Handles issues with pd.ExcelFile or pd.read_excel if file is corrupted/not excel
            messagebox.showerror("File Error", f"Error reading Excel file. It might be corrupted or not a valid Excel file.\nDetails: {ve}")
            self.status_log(f"Error: Could not read Excel file '{file_path}'. Details: {ve}")
            return None, None, []
        except Exception as e:
            import traceback
            messagebox.showerror("Read Error", f"An unexpected error occurred while reading the Excel file: {e}\n\n{traceback.format_exc()}")
            self.status_log(f"Critical Error reading Excel: {e}\n{traceback.format_exc()}")
            return None, None, []

    def get_fcbl_load(self, selected_month_dt): # Removed availability_df from params, uses self.df_availability
        """Fetches FCBL for the first hour of the selected month from self.df_availability."""
        self.status_log(f"Attempting to fetch FCBL for {selected_month_dt.strftime('%b-%y')}.")

        if self.df_availability is None or self.df_availability.empty:
            self.root.after(0, self.status_log_safe, "FCBL fetch failed: Availability data is not loaded or is empty.")
            self.root.after(0, self._show_messagebox_safe, "error", "Data Error", "Cannot fetch FCBL: 'Availability' data is missing or empty.")
            return None

        try:
            # Ensure the relevant columns exist and are of the correct type.
            # Timestamp in 1st col (idx 0), FCBL Load in 4th col (idx 3, Column D)
            if len(self.df_availability.columns) < 4:
                msg = "Availability sheet must have at least four columns (expected Timestamp in Col A, Load in Col D)."
                self.root.after(0, self.status_log_safe, f"FCBL fetch error: {msg}")
                self.root.after(0, self._show_messagebox_safe, "error", "Data Error", msg)
                return None

            # Convert timestamp column (assumed Col A, index 0) if not already datetime
            if not pd.api.types.is_datetime64_any_dtype(self.df_availability.iloc[:, 0]):
                self.root.after(0, self.status_log_safe, "First column of 'Availability' (for FCBL) is not datetime. Attempting conversion...")
                try:
                    self.df_availability.iloc[:, 0] = pd.to_datetime(self.df_availability.iloc[:, 0], errors='coerce')
                    if self.df_availability.iloc[:, 0].isnull().all() and not self.df_availability.empty:
                         self.root.after(0, self._show_messagebox_safe, "error", "Data Error", "First column of 'Availability' (expected dates for FCBL) could not be converted to dates/timestamps.")
                         self.root.after(0, self.status_log_safe, "Error: All values in first column of 'Availability' failed datetime conversion for FCBL.")
                         return None
                    self.root.after(0, self.status_log_safe, "Successfully converted first column of 'Availability' to datetime for FCBL.")
                except Exception as e_conv_avail:
                    self.root.after(0, self._show_messagebox_safe, "error", "Data Error", f"Could not convert first column of 'Availability' to dates/timestamps for FCBL.\nError: {e_conv_avail}")
                    self.root.after(0, self.status_log_safe, f"Error converting first column of 'Availability' for FCBL: {e_conv_avail}")
                    return None

            # Convert load column (assumed Col D, index 3) if not already numeric
            if not pd.api.types.is_numeric_dtype(self.df_availability.iloc[:, 3]):
                self.root.after(0, self.status_log_safe, "Fourth column (Column D) of 'Availability' (for FCBL) is not numeric. Attempting conversion...")
                try:
                    self.df_availability.iloc[:, 3] = pd.to_numeric(self.df_availability.iloc[:, 3], errors='coerce')
                    # Check if all values became NaN after conversion
                    if self.df_availability.iloc[:, 3].isnull().all() and not self.df_availability.empty and not self.df_availability.iloc[:,3].isna().all():
                         self.root.after(0, self._show_messagebox_safe, "error", "Data Error", "Fourth column (Column D) of 'Availability' (expected load values for FCBL) could not be converted to numbers.")
                         self.root.after(0, self.status_log_safe, "Error: All values in fourth column (Column D) of 'Availability' failed numeric conversion for FCBL.")
                         return None
                    self.root.after(0, self.status_log_safe, "Successfully converted fourth column (Column D) of 'Availability' to numeric for FCBL.")
                except Exception as e_conv_load_avail:
                    self.root.after(0, self._show_messagebox_safe, "error", "Data Error", f"Could not convert fourth column (Column D) of 'Availability' (load values) to numbers for FCBL.\nError: {e_conv_load_avail}")
                    self.root.after(0, self.status_log_safe, f"Error converting fourth column (Column D) of 'Availability' for FCBL: {e_conv_load_avail}")
                    return None

            # Define the first hour of the selected month
            start_of_month = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            # FCBL is the 1st hour availability, meaning from 00:00 up to (but not including) 01:00
            end_of_first_hour = start_of_month + datetime.timedelta(hours=1)

            self.root.after(0, self.status_log_safe, f"Filtering Availability data for FCBL between {start_of_month} and {end_of_first_hour}.")

            # Filter for the first hour (00:00 to 00:59 inclusive for that day)
            first_hour_availability_df = self.df_availability[
                (self.df_availability.iloc[:, 0] >= start_of_month) &
                (self.df_availability.iloc[:, 0] < end_of_first_hour) # Use < to exclude 01:00:00
            ]

            if first_hour_availability_df.empty:
                self.root.after(0, self.status_log_safe, f"No availability data found for the first hour (00:00-00:59) of {selected_month_dt.strftime('%b-%y')}.")
                self.root.after(0, self._show_messagebox_safe, "warning", "Data Warning", f"No availability data for the first hour of {selected_month_dt.strftime('%b-%y')}. Cannot determine FCBL.")
                return None

            # Get the first non-NaN availability value in that hour from the load column (iloc[:, 3] for Col D)
            fcbl_value_series = first_hour_availability_df.iloc[:, 3].dropna()
            if fcbl_value_series.empty:
                self.root.after(0, self.status_log_safe, f"All availability values in the first hour of {selected_month_dt.strftime('%b-%y')} (Column D) are NaN.")
                self.root.after(0, self._show_messagebox_safe, "warning", "Data Warning", f"All FCBL values for {selected_month_dt.strftime('%b-%y')} in the first hour are missing (NaN).")
                return None

            fcbl_value = fcbl_value_series.iloc[0] # Get the first valid numeric value
            self.root.after(0, self.status_log_safe, f"FCBL load determined: {fcbl_value}")
            return float(fcbl_value)

        except Exception as e:
            import traceback
            self.root.after(0, self.status_log_safe, f"Error determining FCBL: {e}\n{traceback.format_exc()}")
            self.root.after(0, self._show_messagebox_safe, "error", "FCBL Error", f"Could not determine FCBL load from 'Availability' sheet.\nCheck data format and content.\nError: {e}")
            return None

    def perform_minute_wise_processing(self, dispatch_df_month, availability_df_month, selected_month_dt, start_load):
        """
        Performs the minute-by-minute load calculation for the selected month.

        Args:
            dispatch_df_month (pd.DataFrame): Filtered dispatch instructions for the selected month.
                                              Assumed columns (by iloc):
                                                0: Timestamp (datetime)
                                                1: Notification Type (str, e.g., "RAMP", "TARGET") - This is an interpretation.
                                                2: Value (float, e.g., Ramp Rate in MW/min or Target Load in MW)
            availability_df_month (pd.DataFrame): Filtered availability data for the selected month.
                                                 Assumed columns (by iloc):
                                                   0: Timestamp (datetime)
                                                   1: Available Load (float, MW) - "Final Availability" source.
            selected_month_dt (datetime.datetime): The first day of the month to process.
            start_load (float): The initial load at the beginning of the month.

        Returns:
            pd.DataFrame: DataFrame with 'Date Time Stamp', 'Load', 'Load Per Minute', 'LPM (30 Min Sum)'.
        """
        self.status_log(f"Starting minute-wise processing for {selected_month_dt.strftime('%b-%y')} with initial load: {start_load:.2f} MW.")

        # Determine the number of days and total minutes in the selected month
        month_start_dt = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        next_month_start_dt = (month_start_dt + relativedelta(months=1))
        # num_days = (next_month_start_dt - datetime.timedelta(days=1)).day # Alternative way to get num_days
        total_minutes_in_month = int((next_month_start_dt - month_start_dt).total_seconds() / 60)

        self.status_log(f"Month: {selected_month_dt.strftime('%b-%y')}, Total minutes for processing: {total_minutes_in_month}")

        # Create a complete minute-wise timestamp range for the selected month
        timestamps = pd.date_range(start=month_start_dt, periods=total_minutes_in_month, freq='T')

        results_list = []
        current_load = float(start_load)

        # --- Prepare Dispatch Instructions ---
        # Standardize column access for dispatch instructions if possible, or rely on iloc with clear assumptions.
        # For this implementation, we'll assume:
        # Col 0 (Timestamp): Already datetime.
        # Col 1 (Instruction_Type): String like "RAMP", "TARGET". This is an interpretation.
        # Col 2 (Value): Numeric. Ramp rate (MW/min) if type is "RAMP", Target load (MW) if type is "TARGET".

        # Convert relevant dispatch columns to numeric, coercing errors, and ensure types.
        # This should ideally be done once after reading/filtering.
        # For robustness, let's ensure columns exist and attempt conversion if needed.
        # We need to define what columns are expected for 'Notification Type' and 'Value'.
        # The problem description is not explicit on exact column names for these.
        # Let's assume Column B (index 1) is "Notification Type" (text) and Column C (index 2) is "Value" (numeric).

        dispatch_instructions = []
        if dispatch_df_month is not None and not dispatch_df_month.empty:
            if len(dispatch_df_month.columns) > 2: # Need at least 3 columns (Time, Type, Value)
                # Ensure timestamp is sorted
                dispatch_df_month = dispatch_df_month.sort_values(by=dispatch_df_month.columns[0]).copy()

                # Attempt to make Value column numeric
                try:
                    dispatch_df_month.iloc[:, 2] = pd.to_numeric(dispatch_df_month.iloc[:, 2], errors='coerce')
                except Exception: # Fallback if conversion fails broadly
                    self.status_log("Warning: Could not ensure 'Value' column (assumed 3rd) in Dispatch Instructions is numeric.")

                for _, row in dispatch_df_month.iterrows():
                    instr_time = row.iloc[0]
                    instr_type = str(row.iloc[1]).upper().strip() if pd.notna(row.iloc[1]) else "UNKNOWN"
                    instr_value = row.iloc[2] if pd.notna(row.iloc[2]) else None # Value can be NaN
                    dispatch_instructions.append({'time': instr_time, 'type': instr_type, 'value': instr_value})
            else:
                self.status_log("Warning: Dispatch Instructions sheet has fewer than 3 columns. Expected Time, Type, Value.")

        # --- Prepare Availability Data (for "Final Availability" lookup) ---
        # Assuming availability_df_month has Timestamp (col 0) and Available Load (col 1)
        # We'll create a series for quick lookup, indexed by time.
        final_availability_series = None
        if availability_df_month is not None and not availability_df_month.empty:
            if len(availability_df_month.columns) > 1: # Ensure at least Timestamp and one Load column
                 # Ensure timestamp is datetime and load is numeric (assuming load is in the second column, index 1, for this general lookup)
                try:
                    if not pd.api.types.is_datetime64_any_dtype(availability_df_month.iloc[:, 0]):
                        availability_df_month.iloc[:, 0] = pd.to_datetime(availability_df_month.iloc[:, 0], errors='coerce')

                    # For "Final Availability", the problem doesn't specify which column.
                    # Let's assume it's the same as FCBL for now (Column D, index 3) if available, otherwise try Column B (index 1).
                    # This part might need clarification if "Final Availability" is different from FCBL column.
                    availability_load_col_idx = 1 # Default to second column (index 1)
                    if len(availability_df_month.columns) > 3: # If Column D exists, prefer it for consistency with FCBL
                        availability_load_col_idx = 3

                    if not pd.api.types.is_numeric_dtype(availability_df_month.iloc[:, availability_load_col_idx]):
                        availability_df_month.iloc[:, availability_load_col_idx] = pd.to_numeric(availability_df_month.iloc[:, availability_load_col_idx], errors='coerce')

                    # Set timestamp as index, handle duplicates by taking the first value.
                    # Drop NaT indices and NaN values in the load column before creating series.
                    avail_copy = availability_df_month.dropna(subset=[availability_df_month.columns[0], availability_df_month.columns[availability_load_col_idx]]).copy()
                    if not avail_copy.empty:
                        avail_copy = avail_copy.set_index(avail_copy.columns[0])
                        if not avail_copy.index.has_duplicates:
                            final_availability_series = avail_copy.iloc[:, availability_load_col_idx -1] # Adjust index because set_index removed one col
                        else: # Handle duplicates, e.g., take mean or first
                            final_availability_series = avail_copy.iloc[:, availability_load_col_idx -1].groupby(avail_copy.index).first()
                            self.status_log("Warning: Duplicate timestamps found in Availability data for Final Availability lookup. Used first entry for each.")
                except Exception as e:
                    self.status_log(f"Warning: Could not process Availability data for Final Availability lookup: {e}")
            else:
                self.status_log("Warning: Availability sheet has fewer than 2 columns. Cannot use for 'Final Availability'.")

        instruction_idx = 0
        active_ramp_rate = None  # MW/minute

        self.status_log("Starting iterative minute-by-minute load calculation...")
        log_update_interval = max(1, total_minutes_in_month // 100) # Update progress roughly 100 times

        for i, ts in enumerate(timestamps):
            # Progress update
            if (i + 1) % log_update_interval == 0 or i == 0:
                progress = ((i + 1) / total_minutes_in_month) * 100
                # This method will be called from a thread, so GUI updates must be scheduled.
                self.root.after(0, self.update_progress_safe, progress)
                self.root.after(0, self.status_log_safe, f"Processing {ts.strftime('%Y-%m-%d %H:%M')} ({i+1}/{total_minutes_in_month}). Load: {current_load:.2f}")


            # --- Process Dispatch Instructions ---
            # Logic: Before first notification: use custom load or FCBL (this is 'start_load', handled initially)
            #        Ramp: use ramp rate (no capping at Target Load)
            #        Between instructions: maintain previous load or get from Final Availability

            new_instruction_this_minute = False
            while instruction_idx < len(dispatch_instructions) and ts >= dispatch_instructions[instruction_idx]['time']:
                instr = dispatch_instructions[instruction_idx]
                instr_time = instr['time']
                instr_type = instr['type']
                instr_value = instr['value']
                new_instruction_this_minute = True # Flag that an instruction is processed for this minute

                log_msg_instr = f"  Minute {ts.strftime('%H:%M')}: Processing instruction from {instr_time.strftime('%H:%M')} - Type: {instr_type}, Value: {instr_value}"
                self.root.after(0, self.status_log_safe, log_msg_instr)


                if instr_type == "RAMP":
                    if instr_value is not None:
                        active_ramp_rate = float(instr_value)
                        self.root.after(0, self.status_log_safe, f"    RAMP instruction: New ramp rate = {active_ramp_rate:.2f} MW/min.")
                    else:
                        self.root.after(0, self.status_log_safe, "    RAMP instruction with no value. Ramp rate unchanged or ignored.")
                elif instr_type == "TARGET": # Assuming "TARGET" means set load immediately
                    if instr_value is not None:
                        current_load = float(instr_value)
                        active_ramp_rate = None # A target instruction typically stops any ongoing ramp
                        self.root.after(0, self.status_log_safe, f"    TARGET instruction: Load set to {current_load:.2f} MW. Ramp stopped.")
                    else:
                        self.root.after(0, self.status_log_safe, "    TARGET instruction with no value. Load unchanged.")
                # Add other instruction types here if necessary
                else:
                    self.root.after(0, self.status_log_safe, f"    Unknown instruction type '{instr_type}' or no value. Ignored.")

                instruction_idx += 1

            # --- Apply Ramp or Maintain Logic ---
            if active_ramp_rate is not None:
                # If a new instruction also occurred this minute that set the load (e.g. TARGET),
                # that takes precedence over applying ramp for *this* minute's start.
                # Ramp applies from this minute *onwards*.
                # The problem: "Ramp: use ramp rate (no capping at Target Load)"
                # If a RAMP instruction was processed, the ramp rate is now active.
                # If a TARGET instruction was processed, current_load is set, and ramp is stopped.
                # So, if active_ramp_rate is NOT None here, it means a RAMP instruction is the latest dominant one.
                current_load += active_ramp_rate # Ramp rate is MW/min, loop is per minute.
                # self.status_log(f"    Ramping. Load at end of minute {ts.strftime('%H:%M')}: {current_load:.2f} MW") # Too verbose
            elif not new_instruction_this_minute:
                # "Between instructions: maintain previous load OR get from Final Availability"
                # If no new instruction this minute and not ramping, try to get from Final Availability.
                if final_availability_series is not None:
                    try:
                        # Try to get exact match for timestamp.
                        # Using reindex with method='ffill' could also work if we want to carry forward last known availability.
                        # For "get from Final Availability", an exact match or nearest (within tolerance) seems more appropriate.
                        available_load_at_ts = final_availability_series.get(ts) # Exact match
                        if available_load_at_ts is not None and pd.notna(available_load_at_ts):
                            current_load = float(available_load_at_ts)
                            # self.status_log(f"    Between instructions. Load set from Final Availability at {ts.strftime('%H:%M')}: {current_load:.2f} MW") # Too verbose
                        # else: maintain previous load (implicitly handled as current_load is not changed)
                    except KeyError:
                        # Timestamp not in availability data, maintain previous load
                        pass
                    except Exception as e_fa:
                        self.root.after(0, self.status_log_safe, f"    Error looking up Final Availability for {ts}: {e_fa}. Maintaining previous load.")
                # else: maintain previous load (implicitly handled)

            # Store results for this minute
            load_per_minute = current_load / 30.0 # As per requirement

            results_list.append({
                'Date Time Stamp': ts,
                'Load': round(current_load, 3),
                'Load Per Minute': round(load_per_minute, 5)
            })

        # self.root.after(0, self.update_progress_safe, 100) # Handled by caller
        self.root.after(0, self.status_log_safe, "Core minute-wise load calculation loop finished.")

        if not results_list:
            self.root.after(0, self.status_log_safe, "Warning: No results generated from minute-wise processing. Returning empty DataFrame.")
            return pd.DataFrame(columns=['Date Time Stamp', 'Load', 'Load Per Minute', 'LPM (30 Min Sum)'])

        results_df = pd.DataFrame(results_list)

        # Calculate LPM (30 Min Sum) - rolling sum for each 30-minute window
        if not results_df.empty and 'Load Per Minute' in results_df.columns:
            self.root.after(0, self.status_log_safe, "Calculating 30-minute rolling sum (LPM)...")
            results_df['LPM (30 Min Sum)'] = results_df['Load Per Minute'].rolling(window=30, min_periods=1).sum().round(5)
            self.root.after(0, self.status_log_safe, "LPM (30 Min Sum) calculation complete.")
        else:
            results_df['LPM (30 Min Sum)'] = pd.NA

        return results_df

    def save_output_excel(self, result_df, export_path):
        """Saves the result DataFrame to an Excel file with formatting using openpyxl."""
        self.root.after(0, self.status_log_safe, f"Attempting to save output to: {export_path}")
        try:
            with pd.ExcelWriter(export_path, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
                # Ensure 'Date Time Stamp' is the first column if it exists for consistent formatting
                # and to make it easier to apply specific datetime formatting.
                if 'Date Time Stamp' in result_df.columns:
                    cols = ['Date Time Stamp'] + [col for col in result_df.columns if col != 'Date Time Stamp']
                    result_df_ordered = result_df[cols]
                else:
                    result_df_ordered = result_df.copy() # Use a copy if no reordering

                result_df_ordered.to_excel(writer, index=False, sheet_name="FADL_Calculation")

                # Access the workbook and worksheet objects for formatting
                workbook = writer.book
                worksheet = writer.sheets["FADL_Calculation"]

                # Freeze headers (first row)
                worksheet.freeze_panes = 'A2' # Freeze row 1

                # Define styles
                header_font = Font(bold=True, color="FFFFFFFF", name='Calibri') # White text
                header_fill = PatternFill(start_color="FF0070C0", end_color="FF0070C0", fill_type="solid") # Dark Blue fill
                center_alignment = Alignment(horizontal="center", vertical="center")

                # Apply styles to header row and auto-adjust column widths
                for col_num, column_title in enumerate(result_df_ordered.columns, 1):
                    cell = worksheet.cell(row=1, column=col_num)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_alignment

                    column_letter = get_column_letter(col_num)

                    # Determine max length for column width calculation
                    header_len = len(str(column_title))

                    # For data length, sample or full scan. Using full scan for accuracy.
                    # Convert all data in column to string to find max length safely.
                    # NaNs should be handled or they might cause issues with len().
                    if not result_df_ordered[column_title].empty:
                        max_data_len = result_df_ordered[column_title].astype(str).map(len).max()
                    else:
                        max_data_len = 0

                    max_len = max(header_len, max_data_len)

                    if column_title == 'Date Time Stamp':
                        # Specific width for datetime stamps for consistency
                        adjusted_width = 22
                    else:
                        adjusted_width = (max_len + 2) * 1.1 # Add padding and a bit extra

                    worksheet.column_dimensions[column_letter].width = min(max(adjusted_width, 10), 50) # Min width 10, Max width 50

                # Apply number formats for numeric columns (e.g., Load, Load Per Minute, LPM Sum)
                # And specific datetime format for the 'Date Time Stamp' column
                for col_idx, col_name in enumerate(result_df_ordered.columns):
                    col_letter = get_column_letter(col_idx + 1)
                    if col_name == 'Date Time Stamp':
                        # The datetime_format in ExcelWriter should handle this, but explicit format can be set too.
                        # worksheet.column_dimensions[col_letter].number_format = 'YYYY-MM-DD HH:MM:SS' # Redundant if ExcelWriter handles it
                        pass # pd.ExcelWriter's datetime_format handles this
                    elif pd.api.types.is_numeric_dtype(result_df_ordered[col_name]):
                        # Apply a general number format, e.g., up to 3 decimal places
                        # You can customize this based on expected precision for each column type
                        num_format = '#,##0.000' # Example: 1,234.567 or 0.123
                        if "Per Minute" in col_name or "Sum" in col_name : # Higher precision for these
                            num_format = '#,##0.00000'
                        elif "Load" == col_name: # Standard load precision
                             num_format = '#,##0.00'

                        for row_idx in range(2, worksheet.max_row + 1): # Start from row 2 (data rows)
                            worksheet[f'{col_letter}{row_idx}'].number_format = num_format

            self.root.after(0, self.status_log_safe, f"Output successfully saved and formatted: {export_path}")
            # Schedule messagebox to be shown from main thread
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Processing complete. Output saved to:\n{export_path}"))
        except PermissionError:
            err_msg = f"Permission denied. Failed to save output file to {export_path}.\nPlease ensure the application has write permissions to the directory and the file is not open elsewhere."
            self.root.after(0, self.status_log_safe, err_msg)
            self.root.after(0, lambda: messagebox.showerror("Save Error", err_msg))
        except Exception as e:
            import traceback
            err_msg = f"Failed to save output file.\nError: {e}" # Removed traceback from user-facing message for brevity
            self.root.after(0, self.status_log_safe, f"Error saving output Excel: {e}\n{traceback.format_exc()}") # Log full traceback
            self.root.after(0, lambda: messagebox.showerror("Save Error", err_msg))

    # --- Helper methods for thread-safe GUI updates ---
    def update_progress_safe(self, value):
        """Thread-safe way to update the progress bar."""
        self.progress_bar['value'] = value
        # self.root.update_idletasks() # Might not be needed if called via root.after

    def status_log_safe(self, message):
        """Thread-safe way to add a message to the status log."""
        # This method will be called by self.root.after(0, ...) so it runs in the main GUI thread.
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        # print(message) # Optional: keep console log for debugging from thread

    def _show_messagebox_safe(self, msg_type, title, message):
        """Thread-safe way to show a messagebox. msg_type is 'info', 'warning', or 'error'."""
        # This method will be called by self.root.after(0, ...)
        if msg_type == "info":
            messagebox.showinfo(title, message)
        elif msg_type == "warning":
            messagebox.showwarning(title, message)
        elif msg_type == "error":
            messagebox.showerror(title, message)

    def _processing_logic(self):
        """The actual processing logic that will run in a separate thread."""
        import traceback # Import for detailed error logging in thread
        try:
            self.root.after(0, self.status_log_safe, "Background processing thread started.")

            # Get values from Tkinter variables. This is safe as it's read-only from the thread.
            selected_month_str = self.month_var.get()
            custom_load_val = self.custom_load_var.get() # DoubleVar, already float
            export_dir = self.export_dir_var.get()
            start_load_type = self.start_load_type_var.get()

            # Convert selected month string to datetime object
            try:
                selected_month_dt = datetime.datetime.strptime(selected_month_str, '%b-%y')
            except ValueError:
                # Schedule messagebox and log update on main thread
                self.root.after(0, self._show_messagebox_safe, "error", "Date Error", f"Invalid month format: {selected_month_str}.")
                self.root.after(0, self.status_log_safe, f"Processing aborted in thread: Invalid month format '{selected_month_str}'.")
                return # Exit thread logic

            # --- Determine Start Load ---
            start_load = 0.0
            if start_load_type == "FCBL":
                # get_fcbl_load might show its own message boxes. If it does, it needs to do so via self.root.after.
                # Let's assume get_fcbl_load is modified to do this or returns a clear signal.
                fcbl_val = self.get_fcbl_load(selected_month_dt)
                if fcbl_val is None:
                    # Assuming get_fcbl_load already showed an error message safely.
                    self.root.after(0, self.status_log_safe, "Processing aborted in thread: Could not determine FCBL.")
                    return
                start_load = fcbl_val
            else: # Custom load type
                start_load = custom_load_val # Already validated by GUI thread before starting this thread.

            self.root.after(0, self.status_log_safe, f"Determined start load: {start_load:.2f} MW")

            # --- Filter DataFrames for the selected month ---
            # Accessing self.df_dispatch and self.df_availability (pandas DFs) from thread is generally safe for read operations.
            # .copy() is used to avoid SettingWithCopyWarning if perform_minute_wise_processing modifies them (it shouldn't for inputs).
            month_start_filter = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            month_end_filter = (month_start_filter + relativedelta(months=1))

            # Defensive check and conversion for datetime columns (should be handled by read_excel_data)
            if self.df_dispatch is not None and not self.df_dispatch.empty:
                if not pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                    self.root.after(0, self.status_log_safe, "Warning: Re-converting Dispatch timestamp column in thread.")
                    self.df_dispatch.iloc[:, 0] = pd.to_datetime(self.df_dispatch.iloc[:, 0], errors='coerce')
                dispatch_df_month = self.df_dispatch[
                    (self.df_dispatch.iloc[:, 0] >= month_start_filter) &
                    (self.df_dispatch.iloc[:, 0] < month_end_filter)
                ].copy()
            else: # Should not happen if GUI validation is correct
                dispatch_df_month = pd.DataFrame() # Empty DF

            if self.df_availability is not None and not self.df_availability.empty:
                if not pd.api.types.is_datetime64_any_dtype(self.df_availability.iloc[:, 0]):
                    self.root.after(0, self.status_log_safe, "Warning: Re-converting Availability timestamp column in thread.")
                    self.df_availability.iloc[:, 0] = pd.to_datetime(self.df_availability.iloc[:, 0], errors='coerce')
                availability_df_month = self.df_availability[
                    (self.df_availability.iloc[:, 0] >= month_start_filter) &
                    (self.df_availability.iloc[:, 0] < month_end_filter)
                ].copy()
            else: # Can happen if availability sheet is empty
                availability_df_month = pd.DataFrame()


            self.root.after(0, self.status_log_safe, f"Filtered Dispatch Data for {selected_month_str}: {len(dispatch_df_month)} rows.")
            self.root.after(0, self.status_log_safe, f"Filtered Availability Data for {selected_month_str}: {len(availability_df_month)} rows.")

            if dispatch_df_month.empty:
                 self.root.after(0, self.status_log_safe, f"Warning: No dispatch instructions found for {selected_month_str}. Processing will use start load and availability logic.")

            # --- Perform Core Processing ---
            # This is the CPU-intensive part.
            # perform_minute_wise_processing itself should call self.root.after(0, self.status_log_safe, ...) for its internal logs
            # and self.root.after(0, self.update_progress_safe, ...) for progress bar.
            result_df = self.perform_minute_wise_processing(
                dispatch_df_month,
                availability_df_month,
                selected_month_dt,
                start_load
            )

            # --- Save Output ---
            output_file_name = "FADL Calculation.xlsx"
            import os
            full_export_path = os.path.join(export_dir, output_file_name)

            # save_output_excel also needs to be thread-safe for its GUI interactions (messagebox)
            # It should use self.root.after(0, self._show_messagebox_safe, ...)
            self.save_output_excel(result_df, full_export_path)
            self.root.after(0, self.status_log_safe, "Processing thread finished successfully.")

        except Exception as e:
            error_message = f"Critical error in processing thread: {e}\n{traceback.format_exc()}"
            self.root.after(0, self.status_log_safe, error_message)
            self.root.after(0, self._show_messagebox_safe, "error", "Processing Error", f"An critical unexpected error occurred in the background task: {e}")
        finally:
            # Ensure the button is re-enabled and progress bar reset, always, from main thread
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.progress_bar.config(value=0))
            self.root.after(0, self.status_log_safe, "Background processing thread has ended.")

    def start_processing_thread(self):
        self.status_log("Start processing button clicked. Validating inputs...")

        # --- GUI-Thread Validations (Quick Checks) ---
        if self.df_dispatch is None or self.df_availability is None:
            messagebox.showerror("Input Error", "Please select and successfully load an input Excel file first.")
            self.status_log("Processing aborted: Input data not loaded.")
            return

        selected_month_str = self.month_var.get()
        if not selected_month_str or selected_month_str in ["Select a file first", "Error reading file", "No months found"]:
            messagebox.showerror("Input Error", "Please select a valid month to process.")
            self.status_log("Processing aborted: Month not selected or invalid.")
            return

        if self.start_load_type_var.get() == "Custom":
            try:
                custom_load_val = self.custom_load_var.get() # tk.DoubleVar handles float conversion
                if custom_load_val < 0:
                    messagebox.showerror("Input Error", "Custom load cannot be negative.")
                    self.status_log("Processing aborted: Negative custom load.")
                    return
            except tk.TclError: # Handles if the entry is not a valid float for DoubleVar
                messagebox.showerror("Input Error", "Invalid custom load value. Please enter a number.")
                self.status_log("Processing aborted: Invalid custom load value format.")
                return

        export_dir_path = self.export_dir_var.get()
        if not export_dir_path:
            messagebox.showerror("Input Error", "Please select an export directory.")
            self.status_log("Processing aborted: Export directory not selected.")
            return

        import os # For path validation
        if not os.path.isdir(export_dir_path):
            messagebox.showerror("Input Error", f"The selected export path is not a valid directory:\n{export_dir_path}")
            self.status_log(f"Processing aborted: Export path '{export_dir_path}' is not a directory.")
            return

        self.status_log("Inputs validated by GUI thread. Starting background processing task...")
        self.process_button.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.root.update_idletasks()

        # Run the core processing logic in a separate thread
        # daemon=True means the thread will exit when the main program exits
        processing_thread = threading.Thread(target=self._processing_logic, daemon=True)
        processing_thread.start()

    def status_log(self, message):
        # This method is called from the main GUI thread or scheduled by self.root.after()
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        print(message) # Keep console log for now, good for debugging.

if __name__ == "__main__":
    root = tk.Tk()
    app = LoadProcessorApp(root)
    root.mainloop()
