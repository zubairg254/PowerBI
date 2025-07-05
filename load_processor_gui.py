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
        self.root.title("FADL Load Processor - New Logic") # Updated title
        self.root.geometry("800x700")
        self.root.minsize(600, 500)

        # --- Variables ---
        self.file_path_var = tk.StringVar()
        self.month_var = tk.StringVar()
        # Default to new "FinalAvailabilityHourly" option
        self.start_load_type_var = tk.StringVar(value="FinalAvailabilityHourly")
        self.custom_load_var = tk.DoubleVar(value=0.0)
        self.export_dir_var = tk.StringVar()
        self.df_dispatch = None
        self.df_availability = None
        self.availability_series_hourly_lookup = None # For faster hourly lookups

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

        # --- Starting Load Type ---
        ttk.Label(config_frame, text="Starting Load Type:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)

        # Option 1: Use Final Availability (hourly)
        final_avail_radio = ttk.Radiobutton(config_frame,
                                            text="Use Final Availability (hourly from Col D)",
                                            variable=self.start_load_type_var,
                                            value="FinalAvailabilityHourly",
                                            command=self.toggle_custom_load_entry)
        final_avail_radio.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        # Option 2: Custom Load (Radio button and Entry field)
        custom_load_frame = ttk.Frame(config_frame) # Frame to group "Custom Load" radio and its entry
        custom_load_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=0, pady=0) # Placed below the other radio

        custom_radio = ttk.Radiobutton(custom_load_frame,
                                       text="Custom Load:",
                                       variable=self.start_load_type_var,
                                       value="Custom",
                                       command=self.toggle_custom_load_entry)
        custom_radio.grid(row=0, column=0, sticky=tk.W, pady=2)

        self.custom_load_entry = ttk.Entry(custom_load_frame,
                                           textvariable=self.custom_load_var,
                                           width=10,
                                           state=tk.DISABLED) # Initial state based on default start_load_type_var
        self.custom_load_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        # --- Export Location ---
        export_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        export_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5) # This will be to the right of config
        export_frame.columnconfigure(1, weight=1)

        ttk.Label(export_frame, text="Export Folder:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(export_frame, textvariable=self.export_dir_var, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(export_frame, text="Browse...", command=self.browse_export_dir).grid(row=0, column=2, sticky=tk.E, padx=5, pady=5)

        # --- Controls & Progress ---
        run_frame = ttk.Frame(main_frame, padding="10")
        run_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        run_frame.columnconfigure(0, weight=1)

        self.process_button = ttk.Button(run_frame, text="Start Processing", command=self.start_processing_thread, width=20)
        self.process_button.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(run_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

        # --- Status Log ---
        log_frame = ttk.LabelFrame(main_frame, text="Status Log", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

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
            self.df_dispatch = None
            self.df_availability = None
            self.availability_series_hourly_lookup = None # Reset this too
            self.process_button.config(state=tk.DISABLED)

            _, _, available_months = self.read_excel_data(file_path)
            if self.df_dispatch is not None and available_months:
                self.month_combo['values'] = available_months
                if available_months:
                    self.month_var.set(available_months[0])
                else:
                    self.month_var.set("")
                    self.month_combo['values'] = ["No months found"]
                self.status_log(f"Available months populated: {available_months if available_months else 'None'}")
                if self.df_availability is not None :
                    self.process_button.config(state=tk.NORMAL)
                    # Pre-process availability data for hourly lookup if successfully loaded
                    self._prepare_hourly_availability_lookup()
            else:
                self.month_combo['values'] = ["Error reading/parsing file"]
                self.month_var.set("")
                self.status_log("Could not populate months. Error reading file, missing sheets, or no valid date data in Dispatch Instructions.")
        else:
            self.status_log("File selection cancelled.")

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
            self.custom_load_var.set(0.0)
        self.status_log(f"Start load type set to: {self.start_load_type_var.get()}")

    def _prepare_hourly_availability_lookup(self):
        """ Pre-processes self.df_availability to create a series for faster hourly lookups. """
        self.availability_series_hourly_lookup = None # Reset
        if self.df_availability is None or self.df_availability.empty:
            self.status_log_safe("Availability data is empty, cannot prepare hourly lookup.")
            return

        try:
            # Ensure Timestamp (Col A, idx 0) and Final Availability (Col D, idx 3) exist and are correct types
            if len(self.df_availability.columns) < 4:
                self.status_log_safe("Availability sheet has < 4 columns. Cannot prepare hourly lookup from Col D.")
                return

            # Make a copy for safe modification
            avail_df = self.df_availability.copy()

            # Timestamp conversion (Col A)
            if not pd.api.types.is_datetime64_any_dtype(avail_df.iloc[:, 0]):
                avail_df.iloc[:, 0] = pd.to_datetime(avail_df.iloc[:, 0], errors='coerce')

            # Load value conversion (Col D)
            if not pd.api.types.is_numeric_dtype(avail_df.iloc[:, 3]):
                avail_df.iloc[:, 3] = pd.to_numeric(avail_df.iloc[:, 3], errors='coerce')

            # Drop rows where essential data (timestamp or load for lookup) is NaT/NaN
            avail_df.dropna(subset=[avail_df.columns[0], avail_df.columns[3]], inplace=True)

            if avail_df.empty:
                self.status_log_safe("No valid data in Availability sheet (Cols A & D) for hourly lookup.")
                return

            # Set timestamp as index
            avail_df = avail_df.set_index(avail_df.columns[0])

            # Create a series for the load (Col D data, which is now original index 3, so data index 2 after set_index)
            # This assumes original Col A was index 0, Col B index 1, Col C index 2, Col D index 3.
            # After set_index(ColA), data from Col D is at iloc[:,2] if B,C,D were the only other cols.
            # More robust: use column names if possible, or be very specific about iloc.
            # For now, assuming Col D is the 4th column:
            load_series = avail_df.iloc[:, 2] # This should be data from original Column D

            # Resample to hourly, taking the first value in each hour.
            # This creates an index of hourly timestamps.
            self.availability_series_hourly_lookup = load_series.resample('H').first()
            self.status_log_safe("Hourly availability lookup table prepared.")

        except Exception as e:
            self.status_log_safe(f"Error preparing hourly availability lookup: {e}")
            self.availability_series_hourly_lookup = None


    def get_hourly_final_availability(self, timestamp_obj, availability_df_month=None):
        """
        Looks up final availability from pre-processed series or directly from df_availability (Col D)
        for the given timestamp's hour.
        Uses self.availability_series_hourly_lookup if available (prepared from the whole month's data).
        """
        if self.availability_series_hourly_lookup is None:
            # Fallback or error if lookup not prepared - this should ideally not happen if browse_file worked.
            # For safety, could try a dynamic lookup on availability_df_month if provided.
            # However, the design is to use the pre-processed series.
            self.root.after(0, self.status_log_safe, f"Warning: Hourly availability lookup series not prepared. Cannot get value for {timestamp_obj}.")
            return None

        try:
            # Get the start of the hour for the given timestamp
            lookup_hour = timestamp_obj.replace(minute=0, second=0, microsecond=0)

            # Lookup in the pre-sampled series
            available_load = self.availability_series_hourly_lookup.get(lookup_hour)

            if pd.notna(available_load):
                # self.root.after(0, self.status_log_safe, f"Hourly availability for {lookup_hour}: {available_load}") # Can be too verbose
                return float(available_load)
            else:
                # self.root.after(0, self.status_log_safe, f"No hourly availability found for {lookup_hour} in pre-processed series.") # Can be too verbose
                return None
        except Exception as e:
            self.root.after(0, self.status_log_safe, f"Error during hourly availability lookup for {timestamp_obj}: {e}")
            return None


    def read_excel_data(self, file_path):
        self.status_log(f"Reading Excel file: {file_path}")
        self.df_dispatch = None
        self.df_availability = None
        available_months = []

        if not file_path:
            messagebox.showerror("Error", "No file selected to read.")
            self.status_log("File reading skipped: No file path provided.")
            return None, None, []

        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names

            required_sheets = {"Dispatch Instructions", "Availability"}
            if not required_sheets.issubset(sheet_names):
                missing_sheets = required_sheets - set(sheet_names)
                err_msg = f"Missing required sheet(s): {', '.join(missing_sheets)}."
                messagebox.showerror("Sheet Error", err_msg + f"\nFound sheets: {sheet_names}")
                self.status_log(f"Error: {err_msg} Excel file only contains: {sheet_names}")
                return None, None, []

            self.status_log(f"Found sheets: {sheet_names}. Reading 'Dispatch Instructions' and 'Availability'.")

            self.df_dispatch = pd.read_excel(xls, sheet_name="Dispatch Instructions")
            self.status_log(f"'Dispatch Instructions' sheet read. Rows: {len(self.df_dispatch)}, Columns: {len(self.df_dispatch.columns) if self.df_dispatch is not None else 0}")

            if self.df_dispatch is None or self.df_dispatch.empty:
                self.status_log("Warning: 'Dispatch Instructions' sheet is empty or failed to load. No months to process.")
                self.df_dispatch = None
            elif self.df_dispatch.columns.empty:
                 self.status_log("Warning: 'Dispatch Instructions' sheet has no columns.")
                 messagebox.showwarning("Data Warning", "'Dispatch Instructions' sheet has no columns.")
                 self.df_dispatch = None
            elif not pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]): # Check Col A for dates
                self.status_log("First column of 'Dispatch Instructions' is not datetime. Attempting conversion...")
                try:
                    self.df_dispatch.iloc[:, 0] = pd.to_datetime(self.df_dispatch.iloc[:, 0], errors='coerce')
                    if self.df_dispatch.iloc[:, 0].isnull().all() and not self.df_dispatch.empty:
                         messagebox.showerror("Data Error", "First column of 'Dispatch Instructions' (expected dates) could not be converted to dates/timestamps. All values are invalid.")
                         self.status_log("Error: All values in first column of 'Dispatch Instructions' failed datetime conversion.")
                         self.df_dispatch = None
                    elif self.df_dispatch.iloc[:, 0].isnull().any():
                        num_failed = self.df_dispatch.iloc[:,0].isnull().sum()
                        self.status_log(f"Warning: {num_failed} values in the first column of 'Dispatch Instructions' failed datetime conversion and were set to NaT.")
                        messagebox.showwarning("Data Conversion Warning", f"{num_failed} date entries in 'Dispatch Instructions' are invalid and were ignored.")
                    # else: self.status_log("Successfully converted first column of 'Dispatch Instructions' to datetime.") # Already logged if successful by caller
                except Exception as e_conv:
                    messagebox.showerror("Data Error", f"Could not convert first column of 'Dispatch Instructions' to dates/timestamps.\nError: {e_conv}")
                    self.status_log(f"Error converting first column of 'Dispatch Instructions': {e_conv}")
                    self.df_dispatch = None

            if self.df_dispatch is not None and not self.df_dispatch.empty and pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                valid_dates = self.df_dispatch.iloc[:, 0].dropna()
                if not valid_dates.empty:
                    available_months = sorted(valid_dates.dt.strftime('%b-%y').unique().tolist())
                if not available_months:
                    self.status_log("No valid months found in 'Dispatch Instructions' (Column A after processing).")

            self.df_availability = pd.read_excel(xls, sheet_name="Availability")
            self.status_log(f"'Availability' sheet read. Rows: {len(self.df_availability)}, Columns: {len(self.df_availability.columns) if self.df_availability is not None else 0}")

            if self.df_availability is None or self.df_availability.empty:
                 self.status_log("Warning: 'Availability' sheet is empty or failed to load.")
                 self.df_availability = None
            elif self.df_availability.columns.empty:
                self.status_log("Warning: 'Availability' sheet has no columns.")
                messagebox.showwarning("Data Warning", "'Availability' sheet has no columns.")
                self.df_availability = None
            elif len(self.df_availability.columns) < 4:
                 self.status_log("Warning: 'Availability' sheet has fewer than 4 columns. Required for Final Availability lookup from Column D.")
                 messagebox.showwarning("Data Warning", "'Availability' sheet has fewer than 4 columns. Final Availability features might not work.")
                 # Not setting df_availability to None here, as other parts might still be usable or it might be ignorable.

            return self.df_dispatch, self.df_availability, available_months

        except FileNotFoundError:
            messagebox.showerror("File Error", f"File not found: {file_path}")
            self.status_log(f"Error: File not found at {file_path}")
            return None, None, []
        except ValueError as ve:
            messagebox.showerror("File Error", f"Error reading Excel file. It might be corrupted or not a valid Excel file.\nDetails: {ve}")
            self.status_log(f"Error: Could not read Excel file '{file_path}'. Details: {ve}")
            return None, None, []
        except Exception as e:
            import traceback
            messagebox.showerror("Read Error", f"An unexpected error occurred while reading the Excel file: {e}\n\n{traceback.format_exc()}")
            self.status_log(f"Critical Error reading Excel: {e}\n{traceback.format_exc()}")
            return None, None, []

    def get_initial_availability_load(self, selected_month_dt):
        """Fetches the initial load from 'Final Availability' (Col D) for the first hour of the selected month."""
        self.root.after(0, self.status_log_safe, f"Attempting to fetch Initial Hourly Availability for {selected_month_dt.strftime('%b-%y')}.")

        timestamp_first_minute_of_month = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        initial_load = self.get_hourly_final_availability(timestamp_first_minute_of_month)

        if initial_load is not None:
            self.root.after(0, self.status_log_safe, f"Initial Hourly Availability load determined: {initial_load} (from Col D of Availability sheet for hour 00)")
            return initial_load
        else:
            self.root.after(0, self.status_log_safe, f"No availability data found for the first hour (00:00-00:59) of {selected_month_dt.strftime('%b-%y')} in Col D for Initial Load.")
            self.root.after(0, self._show_messagebox_safe, "warning", "Data Warning", f"No availability data for the first hour of {selected_month_dt.strftime('%b-%y')} (Col D). Cannot determine Initial Hourly Availability.")
            return None


    def perform_minute_wise_processing(self, dispatch_df_month, availability_df_month, selected_month_dt, start_load):
        """
        Performs the minute-by-minute load calculation for the selected month using revised logic.

        Dispatch Instructions Columns (by index for robustness, assuming standard naming if available):
            Col A (idx 0): Notification/Instruction Time
            Col B (idx 1): Target Time Stamp (when ramp should conclude/target is met)
            Col C (idx 2): Ramp Duration (minutes)
            Col E (idx 4): Post-Ramp Target Type (e.g., "FCBL")
            Col F (idx 5): Target Demand (MW)

        Availability Sheet Columns:
            Col A (idx 0): Timestamp
            Col D (idx 3): Final Availability (MW)
        """
        self.root.after(0, self.status_log_safe, f"Starting minute-wise processing for {selected_month_dt.strftime('%b-%y')} with initial load: {start_load:.2f} MW.")

        month_start_dt = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        next_month_start_dt = (month_start_dt + relativedelta(months=1))
        total_minutes_in_month = int((next_month_start_dt - month_start_dt).total_seconds() / 60)

        self.root.after(0, self.status_log_safe, f"Month: {selected_month_dt.strftime('%b-%y')}, Total minutes: {total_minutes_in_month}")
        timestamps = pd.date_range(start=month_start_dt, periods=total_minutes_in_month, freq='T')

        results_list = []
        current_load = float(start_load)

        # Prepare dispatch instructions (parse and validate)
        parsed_instructions = []
        if dispatch_df_month is not None and not dispatch_df_month.empty:
            # Ensure required columns are present by trying to access them by index.
            # Minimum 6 columns for A, B, C, (skip D), E, F
            if len(dispatch_df_month.columns) < 6:
                self.root.after(0, self.status_log_safe, "Dispatch Instructions sheet has insufficient columns. Expected at least 6 columns for new logic (A, B, C, E, F).")
            else:
                temp_df = dispatch_df_month.copy()
                # Convert types safely
                try:
                    temp_df.iloc[:, 0] = pd.to_datetime(temp_df.iloc[:, 0], errors='coerce') # Col A: Notification Time
                    temp_df.iloc[:, 1] = pd.to_datetime(temp_df.iloc[:, 1], errors='coerce') # Col B: Target Time Stamp
                    temp_df.iloc[:, 2] = pd.to_numeric(temp_df.iloc[:, 2], errors='coerce') # Col C: Ramp Duration
                    # Col D is skipped
                    temp_df.iloc[:, 4] = temp_df.iloc[:, 4].astype(str).str.upper().str.strip() # Col E: Post-Ramp Target Type
                    temp_df.iloc[:, 5] = pd.to_numeric(temp_df.iloc[:, 5], errors='coerce') # Col F: Target Demand MW

                    # Drop rows where key information is missing after conversion
                    temp_df.dropna(subset=[temp_df.columns[0], temp_df.columns[5]], inplace=True) # Must have Notification Time and Target Demand

                    for _, row in temp_df.sort_values(by=temp_df.columns[0]).iterrows():
                        parsed_instructions.append({
                            'instr_time': row.iloc[0],
                            'target_time_stamp': row.iloc[1] if pd.notna(row.iloc[1]) else None,
                            'ramp_duration_minutes': row.iloc[2] if pd.notna(row.iloc[2]) else 0, # Default to 0 if NaN
                            'post_ramp_target_type': row.iloc[4],
                            'target_demand_mw': row.iloc[5]
                        })
                except IndexError:
                     self.root.after(0, self.status_log_safe, "Error accessing expected columns in Dispatch Instructions during parsing. Check sheet structure.")
                except Exception as e_parse:
                     self.root.after(0, self.status_log_safe, f"Error parsing Dispatch Instructions: {e_parse}")
            if not parsed_instructions and not dispatch_df_month.empty : # If parsing failed but df wasn't empty
                 self.root.after(0, self.status_log_safe, "Warning: Dispatch Instructions found but could not be parsed into usable format. Processing as if no instructions.")


        instr_idx = 0
        active_ramp_rate = 0.0
        ramp_end_time = None
        is_ramping = False
        current_ramp_target_mw = None # Stores the target_demand_mw of the current/last ramp
        current_post_ramp_type = None # Stores the post_ramp_target_type of the current/last ramp

        log_update_interval = max(1, total_minutes_in_month // 100)

        for i, ts in enumerate(timestamps):
            if (i + 1) % log_update_interval == 0 or i == 0:
                progress = ((i + 1) / total_minutes_in_month) * 100
                self.root.after(0, self.update_progress_safe, progress)
                self.root.after(0, self.status_log_safe, f"Processing {ts.strftime('%Y-%m-%d %H:%M')} ({i+1}/{total_minutes_in_month}). Load: {current_load:.2f}")

            acted_on_instruction_this_minute = False

            # Check for new instruction
            if instr_idx < len(parsed_instructions) and ts == parsed_instructions[instr_idx]['instr_time']:
                instr = parsed_instructions[instr_idx]
                self.root.after(0, self.status_log_safe, f"  Minute {ts.strftime('%H:%M')}: Instruction received. Target Demand: {instr['target_demand_mw']:.2f} MW, Duration: {instr['ramp_duration_minutes']} min, Target Time: {instr['target_time_stamp']}, Post-Ramp: {instr['post_ramp_target_type']}")

                acted_on_instruction_this_minute = True
                current_ramp_target_mw = instr['target_demand_mw'] # Store for end of ramp
                current_post_ramp_type = instr['post_ramp_target_type'] # Store for end of ramp

                if pd.notna(instr['ramp_duration_minutes']) and instr['ramp_duration_minutes'] > 0:
                    duration = instr['ramp_duration_minutes']
                    if current_load == current_ramp_target_mw: # Already at target
                        active_ramp_rate = 0.0
                        self.root.after(0, self.status_log_safe, f"    Load {current_load:.2f} already at target {current_ramp_target_mw:.2f}. No ramp needed.")
                    else:
                        active_ramp_rate = (current_ramp_target_mw - current_load) / duration

                    # Determine ramp_end_time: Prefer Col B (target_time_stamp) if valid
                    if pd.notna(instr['target_time_stamp']) and instr['target_time_stamp'] > ts:
                        ramp_end_time = instr['target_time_stamp']
                    else: # Fallback to duration from Col C, or if Col B is not in future
                        ramp_end_time = ts + datetime.timedelta(minutes=duration)
                        if pd.notna(instr['target_time_stamp']) and instr['target_time_stamp'] <= ts:
                             self.root.after(0, self.status_log_safe, f"    Warning: Target Time Stamp {instr['target_time_stamp']} is not after Instruction Time {ts}. Using calculated ramp end: {ramp_end_time}")

                    is_ramping = True
                    self.root.after(0, self.status_log_safe, f"    Starting ramp. Rate: {active_ramp_rate:.2f} MW/min. Expected End: {ramp_end_time.strftime('%Y-%m-%d %H:%M') if ramp_end_time else 'N/A'}")
                else: # Instantaneous change
                    self.root.after(0, self.status_log_safe, f"    Instantaneous load change to {current_ramp_target_mw:.2f} MW.")
                    current_load = current_ramp_target_mw
                    is_ramping = False
                    active_ramp_rate = 0.0
                    # Apply post-ramp logic immediately for instantaneous changes
                    if current_post_ramp_type == "FCBL":
                        hourly_avail = self.get_hourly_final_availability(ts)
                        if hourly_avail is not None:
                            current_load = hourly_avail
                            self.root.after(0, self.status_log_safe, f"    Instantaneous FCBL applied. Load set to Final Availability: {current_load:.2f} MW for hour of {ts.strftime('%H')}:00.")
                        else:
                            self.root.after(0, self.status_log_safe, f"    Instantaneous FCBL requested, but no Final Availability found for hour of {ts.strftime('%H')}:00. Load remains {current_load:.2f} MW.")
                instr_idx += 1

            # Apply ramp if active
            if is_ramping:
                if ts < ramp_end_time: # Check if ramp_end_time is valid
                    current_load += active_ramp_rate
                    # self.root.after(0, self.status_log_safe, f"    Ramping: Load -> {current_load:.2f}") # Too verbose
                else: # Ramp duration is over or target time reached
                    self.root.after(0, self.status_log_safe, f"  Minute {ts.strftime('%H:%M')}: Ramp ended. Reached ramp target load: {current_ramp_target_mw:.2f} MW.")
                    current_load = current_ramp_target_mw # Snap to target
                    is_ramping = False
                    active_ramp_rate = 0.0
                    # Apply post-ramp logic
                    if current_post_ramp_type == "FCBL":
                        hourly_avail = self.get_hourly_final_availability(ts)
                        if hourly_avail is not None:
                            current_load = hourly_avail
                            self.root.after(0, self.status_log_safe, f"    Post-ramp FCBL applied. Load set to Final Availability: {current_load:.2f} MW for hour of {ts.strftime('%H')}:00.")
                        else:
                            self.root.after(0, self.status_log_safe, f"    Post-ramp FCBL requested, but no Final Availability for hour of {ts.strftime('%H')}:00. Load remains {current_load:.2f} MW.")
                    # else: load is already current_ramp_target_mw, which is correct.

            # If no instruction processed this minute and not ramping, load is maintained (no change to current_load)

            results_list.append({
                'Date Time Stamp': ts,
                'Load': round(current_load, 3),
                'Load Per Minute': round(current_load / 30.0, 5)
            })

        self.root.after(0, self.update_progress_safe, 100)
        self.root.after(0, self.status_log_safe, "Core minute-wise load calculation loop finished.")

        if not results_list:
            self.root.after(0, self.status_log_safe, "Warning: No results generated. Returning empty DataFrame.")
            return pd.DataFrame(columns=['Date Time Stamp', 'Load', 'Load Per Minute', 'LPM (30 Min Sum)'])

        results_df = pd.DataFrame(results_list)

        if not results_df.empty and 'Load Per Minute' in results_df.columns:
            self.root.after(0, self.status_log_safe, "Calculating 30-minute rolling sum (LPM)...")
            results_df['LPM (30 Min Sum)'] = results_df['Load Per Minute'].rolling(window=30, min_periods=1).sum().round(5)
            self.root.after(0, self.status_log_safe, "LPM (30 Min Sum) calculation complete.")
        else:
            results_df['LPM (30 Min Sum)'] = pd.NA

        return results_df


    def save_output_excel(self, result_df, export_path):
        self.root.after(0, self.status_log_safe, f"Attempting to save output to: {export_path}")
        try:
            with pd.ExcelWriter(export_path, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
                if 'Date Time Stamp' in result_df.columns:
                    cols = ['Date Time Stamp'] + [col for col in result_df.columns if col != 'Date Time Stamp']
                    result_df_ordered = result_df[cols]
                else:
                    result_df_ordered = result_df.copy()

                result_df_ordered.to_excel(writer, index=False, sheet_name="FADL_Calculation")
                workbook = writer.book
                worksheet = writer.sheets["FADL_Calculation"]
                worksheet.freeze_panes = 'A2'
                header_font = Font(bold=True, color="FFFFFFFF", name='Calibri')
                header_fill = PatternFill(start_color="FF0070C0", end_color="FF0070C0", fill_type="solid")
                center_alignment = Alignment(horizontal="center", vertical="center")

                for col_num, column_title in enumerate(result_df_ordered.columns, 1):
                    cell = worksheet.cell(row=1, column=col_num)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_alignment
                    column_letter = get_column_letter(col_num)
                    header_len = len(str(column_title))
                    if not result_df_ordered[column_title].empty:
                        max_data_len = result_df_ordered[column_title].astype(str).map(len).max()
                    else: max_data_len = 0
                    max_len = max(header_len, max_data_len)
                    if column_title == 'Date Time Stamp': adjusted_width = 22
                    else: adjusted_width = (max_len + 2) * 1.1
                    worksheet.column_dimensions[column_letter].width = min(max(adjusted_width, 10), 50)

                for col_idx, col_name in enumerate(result_df_ordered.columns):
                    col_letter = get_column_letter(col_idx + 1)
                    if col_name != 'Date Time Stamp' and pd.api.types.is_numeric_dtype(result_df_ordered[col_name]):
                        num_format = '#,##0.000'
                        if "Per Minute" in col_name or "Sum" in col_name : num_format = '#,##0.00000'
                        elif "Load" == col_name: num_format = '#,##0.00'
                        for row_idx in range(2, worksheet.max_row + 1):
                            worksheet[f'{col_letter}{row_idx}'].number_format = num_format

            self.root.after(0, self.status_log_safe, f"Output successfully saved and formatted: {export_path}")
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Processing complete. Output saved to:\n{export_path}"))
        except PermissionError:
            err_msg = f"Permission denied to save to {export_path}. Check permissions/if file is open."
            self.root.after(0, self.status_log_safe, err_msg)
            self.root.after(0, lambda: messagebox.showerror("Save Error", err_msg))
        except Exception as e:
            import traceback
            err_msg = f"Failed to save output file.\nError: {e}"
            self.root.after(0, self.status_log_safe, f"Error saving output Excel: {e}\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Save Error", err_msg))

    def update_progress_safe(self, value):
        self.progress_bar['value'] = value

    def status_log_safe(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _show_messagebox_safe(self, msg_type, title, message):
        if msg_type == "info": messagebox.showinfo(title, message)
        elif msg_type == "warning": messagebox.showwarning(title, message)
        elif msg_type == "error": messagebox.showerror(title, message)

    def _processing_logic(self):
        import traceback
        try:
            self.root.after(0, self.status_log_safe, "Background processing thread started.")
            selected_month_str = self.month_var.get()
            custom_load_val = self.custom_load_var.get()
            export_dir = self.export_dir_var.get()
            start_load_type = self.start_load_type_var.get()

            try:
                selected_month_dt = datetime.datetime.strptime(selected_month_str, '%b-%y')
            except ValueError:
                self.root.after(0, self._show_messagebox_safe, "error", "Date Error", f"Invalid month format: {selected_month_str}.")
                self.root.after(0, self.status_log_safe, f"Processing aborted: Invalid month format '{selected_month_str}'.")
                return

            start_load = 0.0
            if start_load_type == "FinalAvailabilityHourly":
                start_load = self.get_initial_availability_load(selected_month_dt)
                if start_load is None:
                    self.root.after(0, self.status_log_safe, "Processing aborted: Could not determine Initial Hourly Availability for starting load.")
                    return
            else: # Custom
                start_load = custom_load_val

            self.root.after(0, self.status_log_safe, f"Determined start load: {start_load:.2f} MW")

            month_start_filter = selected_month_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            month_end_filter = (month_start_filter + relativedelta(months=1))

            dispatch_df_month = pd.DataFrame()
            if self.df_dispatch is not None and not self.df_dispatch.empty:
                if not pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                    self.root.after(0, self.status_log_safe, "Warning: Re-converting Dispatch timestamp column in thread before filtering.")
                    try: self.df_dispatch.iloc[:, 0] = pd.to_datetime(self.df_dispatch.iloc[:, 0], errors='coerce')
                    except: self.root.after(0, self.status_log_safe, "Error during defensive conversion of Dispatch timestamp.")

                if pd.api.types.is_datetime64_any_dtype(self.df_dispatch.iloc[:, 0]):
                    dispatch_df_month = self.df_dispatch[(self.df_dispatch.iloc[:, 0] >= month_start_filter) & (self.df_dispatch.iloc[:, 0] < month_end_filter)].copy()

            availability_df_month = pd.DataFrame() # This is passed to perform_minute_wise_processing but that function uses self.availability_series_hourly_lookup
            # The self.availability_series_hourly_lookup is prepared from the whole self.df_availability in browse_file.
            # So, availability_df_month is not strictly needed by the new perform_minute_wise_processing if it uses the preprocessed series.
            # However, get_hourly_final_availability can be made to accept it as an argument if needed.
            # For now, let's ensure the pre-processed series is ready if the main df_availability exists.
            if self.df_availability is not None and not self.df_availability.empty and self.availability_series_hourly_lookup is None:
                 self.root.after(0, self.status_log_safe, "Warning: Hourly availability lookup was not prepared. Attempting now.")
                 self._prepare_hourly_availability_lookup() # Attempt to prepare it if not done

            self.root.after(0, self.status_log_safe, f"Filtered Dispatch Data for {selected_month_str}: {len(dispatch_df_month)} rows.")
            # No need to log filtered availability_df_month length if using pre-processed series primarily for lookups.

            if dispatch_df_month.empty and start_load_type != "Custom" and (start_load == 0 or start_load is None) :
                 self.root.after(0, self.status_log_safe, f"Warning: No dispatch instructions for {selected_month_str} and starting load is 0 or could not be determined from availability.")

            result_df = self.perform_minute_wise_processing(dispatch_df_month, self.df_availability, selected_month_dt, start_load) # Pass full self.df_availability for get_hourly_final_availability

            output_file_name = "FADL Calculation.xlsx"
            import os
            full_export_path = os.path.join(export_dir, output_file_name)
            self.save_output_excel(result_df, full_export_path)
            self.root.after(0, self.status_log_safe, "Processing thread finished successfully.")

        except Exception as e:
            error_message = f"Critical error in processing thread: {e}\n{traceback.format_exc()}"
            self.root.after(0, self.status_log_safe, error_message)
            self.root.after(0, self._show_messagebox_safe, "error", "Processing Error", f"An critical unexpected error occurred: {e}")
        finally:
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.progress_bar.config(value=0))
            self.root.after(0, self.status_log_safe, "Background processing thread has ended.")

    def start_processing_thread(self):
        self.status_log("Start processing button clicked. Validating inputs...")
        if self.df_dispatch is None or self.df_availability is None:
             # Check specifically if availability is None only if it's needed for start load
            if self.start_load_type_var.get() == "FinalAvailabilityHourly" and self.df_availability is None:
                messagebox.showerror("Input Error", "Availability data not loaded. Please select a valid Excel file with an 'Availability' sheet for this starting load option.")
                self.status_log("Processing aborted: Availability data required but not loaded.")
                return
            elif self.df_dispatch is None: # Dispatch is always needed
                 messagebox.showerror("Input Error", "Dispatch data not loaded. Please select a valid Excel file with a 'Dispatch Instructions' sheet.")
                 self.status_log("Processing aborted: Dispatch data not loaded.")
                 return


        selected_month_str = self.month_var.get()
        if not selected_month_str or selected_month_str in ["Select a file first", "Error reading file", "No months found", "Error in sheet data", "Error reading/parsing file"]:
            messagebox.showerror("Input Error", "Please select a valid month to process from a successfully loaded file.")
            self.status_log("Processing aborted: Month not selected or invalid.")
            return

        if self.start_load_type_var.get() == "Custom":
            try:
                custom_load_val = self.custom_load_var.get()
                if custom_load_val < 0:
                    messagebox.showerror("Input Error", "Custom load cannot be negative.")
                    self.status_log("Processing aborted: Negative custom load.")
                    return
            except tk.TclError:
                messagebox.showerror("Input Error", "Invalid custom load value. Please enter a number.")
                self.status_log("Processing aborted: Invalid custom load value format.")
                return

        export_dir_path = self.export_dir_var.get()
        if not export_dir_path:
            messagebox.showerror("Input Error", "Please select an export directory.")
            self.status_log("Processing aborted: Export directory not selected.")
            return

        import os
        if not os.path.isdir(export_dir_path):
            messagebox.showerror("Input Error", f"The selected export path is not a valid directory:\n{export_dir_path}")
            self.status_log(f"Processing aborted: Export path '{export_dir_path}' is not a directory.")
            return

        self.status_log("Inputs validated. Starting background processing task...")
        self.process_button.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.root.update_idletasks()

        processing_thread = threading.Thread(target=self._processing_logic, daemon=True)
        processing_thread.start()

    def status_log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        print(message)

if __name__ == "__main__":
    root = tk.Tk()
    app = LoadProcessorApp(root)
    root.mainloop()
