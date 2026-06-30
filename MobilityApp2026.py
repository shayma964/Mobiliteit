# -*- coding: utf-8 -*-
"""
Mobility Analysis Tool for UGent

A Tkinter-based GUI application for processing and analyzing UGent employee
mobility data from Excel files. Provides data mapping, analysis execution,
pivot table generation, and result export functionality.

This tool integrates with the UGent Mobility Backend to process transport
mode assignments across multiple data sources (MOB0-MOB6).
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ugent_mobility_backend_fixed import MobilityEngine, apply_mapping, REQUIRED_FIELDS
import pandas as pd
import threading
from datetime import datetime
import os
import traceback
import sys 


def resource_path(relative_path):
    """
    Resolve resource paths for packaged applications.
    
    Returns the correct path for resources whether the application is
    running as a frozen PyInstaller executable or from source.
    
    Args:
        relative_path (str): Relative path to the resource.
        
    Returns:
        str: Absolute path to the resource.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_app_path():
    """
    Get the application base directory.
    
    Returns the directory where the application is located, handling both
    frozen executables and source code execution.
    
    Returns:
        str: Absolute path to the application directory.
    """
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(__file__)


os.chdir(get_app_path())


class App:
    """
    Main Tkinter application for mobility analysis.
    
    Manages the GUI, file handling, data mapping, analysis execution,
    and result visualization for UGent mobility data processing.
    
    Attributes:
        root (tk.Tk): The root Tkinter window.
        file (str): Path to the main Excel file.
        afst (str): Path to the distance reference file.
        df (pd.DataFrame): Processed results dataframe.
        data (dict): Loaded Excel sheet data.
        mapping (dict): Column mapping configuration.
        afstanden_mapping (dict): Distance file column mapping.
        engine (MobilityEngine): The analysis engine instance.
        pivot_table (pd.DataFrame): Pivot table of results.
    """

    def __init__(self, root):
        """
        Initialize the application window and data structures.
        
        Sets up the root window, initializes data variables, and builds
        the user interface.
        
        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.title("Mobility Tool UGent")
        self.root.geometry("1300x800")
        
        self.file = None
        self.afst = None
        self.df = None
        self.data = {}
        self.mapping = {}
        self.afstanden_mapping = {}
        self.engine = None
        self.pivot_table = None
        
        self.build()
        
    def build(self):
        """
        Construct the main user interface.
        
        Creates the control panel with file loading buttons, date inputs,
        analysis controls, and the data preview table with scrollbars.
        """
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        control_frame = tk.LabelFrame(main_frame, text="Control Panel", font=("Arial", 10, "bold"))
        control_frame.pack(fill="x", pady=(0, 10))
        
        file_row = tk.Frame(control_frame)
        file_row.pack(fill="x", padx=10, pady=10)
        
        tk.Button(file_row, text="📁 Load Main Excel", command=self.load_main, 
                 bg="#4CAF50", fg="white", font=("Arial", 10), padx=20, pady=5).pack(side="left", padx=5)
        tk.Button(file_row, text="📏 Load Distance File", command=self.load_afst,
                 bg="#2196F3", fg="white", font=("Arial", 10), padx=20, pady=5).pack(side="left", padx=5)
        
        self.main_status = tk.Label(file_row, text="No main file loaded", fg="gray", font=("Arial", 9))
        self.main_status.pack(side="left", padx=10)
        
        self.dist_status = tk.Label(file_row, text="No distance file loaded", fg="gray", font=("Arial", 9))
        self.dist_status.pack(side="left", padx=10)
        
        date_frame = tk.Frame(control_frame)
        date_frame.pack(fill="x", padx=10, pady=5)

        start_frame = tk.Frame(date_frame)
        start_frame.pack(side="left", padx=5)

        tk.Label(start_frame, text="Start Date:", font=("Arial", 9)).pack(side="left")
        self.start_entry = tk.Entry(start_frame, width=12, font=("Arial", 9))
        self.start_entry.insert(0, "2025-10-01")
        self.start_entry.pack(side="left", padx=5)

        tk.Label(start_frame, text="(YYYY-MM-DD)", font=("Arial", 7), fg="gray").pack(side="left")

        end_frame = tk.Frame(date_frame)
        end_frame.pack(side="left", padx=5)

        tk.Label(end_frame, text="End Date:", font=("Arial", 9)).pack(side="left")
        self.end_entry = tk.Entry(end_frame, width=12, font=("Arial", 9))
        self.end_entry.insert(0, "2025-10-31")
        self.end_entry.pack(side="left", padx=5)

        tk.Label(end_frame, text="(YYYY-MM-DD)", font=("Arial", 7), fg="gray").pack(side="left")

        tk.Button(start_frame, text="📅", command=lambda: self.show_date_picker(self.start_entry),
                 width=2, font=("Arial", 8)).pack(side="left", padx=2)

        tk.Button(end_frame, text="📅", command=lambda: self.show_date_picker(self.end_entry),
                 width=2, font=("Arial", 8)).pack(side="left", padx=2)

        threshold_frame = tk.Frame(control_frame)
        threshold_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(threshold_frame, text="Minimum bike days for 'Fietser':",
         font=("Arial", 9)).pack(side="left", padx=5)
        self.threshold_entry = tk.Entry(threshold_frame, width=5, font=("Arial", 9))
        self.threshold_entry.insert(0, "4")
        self.threshold_entry.pack(side="left", padx=5)
        tk.Label(threshold_frame, text="days", font=("Arial", 9)).pack(side="left")

        action_frame = tk.Frame(control_frame)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        self.run_btn = tk.Button(action_frame, text="▶ RUN ANALYSIS", command=self.run_analysis,
                                 bg="#FF9800", fg="white", font=("Arial", 10, "bold"), padx=30, pady=5)
        self.run_btn.pack(side="left", padx=5)
        
        self.pivot_btn = tk.Button(action_frame, text="📊 Show Table", command=self.show_pivot_table,
                                   bg="#ff6a6a", fg="white", font=("Arial", 10,"bold"), padx=20, pady=5, state="disabled")
        self.pivot_btn.pack(side="left", padx=5)
        
        self.save_btn = tk.Button(action_frame, text="💾 Save Results", command=self.save_results,
                                  bg="#F44336", fg="white", font=("Arial", 10,"bold"), padx=20, pady=5, state="disabled")
        self.save_btn.pack(side="left", padx=5)
        
        progress_frame = tk.Frame(control_frame)
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=300)
        self.progress.pack(side="left", padx=5)
        
        self.status_label = tk.Label(progress_frame, text="Ready", fg="green", font=("Arial", 9))
        self.status_label.pack(side="left", padx=10)
        
        preview_frame = tk.LabelFrame(main_frame, text="Data Preview", font=("Arial", 10, "bold"))
        preview_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        tree_container = tk.Frame(preview_frame)
        tree_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_container)
        scrollbar_y.pack(side="right", fill="y")
        
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        
        self.tree = ttk.Treeview(tree_container, yscrollcommand=scrollbar_y.set, 
                                  xscrollcommand=scrollbar_x.set, height=20)
        self.tree.pack(fill="both", expand=True)
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill="x")
        
        self.info_label = tk.Label(info_frame, text="", fg="blue", font=("Arial", 9))
        self.info_label.pack(side="left")
        end_frame.pack(side="left", padx=5)

        tk.Label(end_frame, text="End Date:", font=("Arial", 9)).pack(side="left")
        self.end_entry = tk.Entry(end_frame, width=12, font=("Arial", 9))
        self.end_entry.insert(0, "2025-10-31")
        self.end_entry.pack(side="left", padx=5)

        tk.Label(end_frame, text="(YYYY-MM-DD)", font=("Arial", 7), fg="gray").pack(side="left")

        tk.Button(start_frame, text="📅", command=lambda: self.show_date_picker(self.start_entry),
                 width=2, font=("Arial", 8)).pack(side="left", padx=2)

        tk.Button(end_frame, text="📅", command=lambda: self.show_date_picker(self.end_entry),
                 width=2, font=("Arial", 8)).pack(side="left", padx=2)

        threshold_frame = tk.Frame(control_frame)
        threshold_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(threshold_frame, text="Minimum bike days for 'Fietser':",
         font=("Arial", 9)).pack(side="left", padx=5)
        self.threshold_entry = tk.Entry(threshold_frame, width=5, font=("Arial", 9))
        self.threshold_entry.insert(0, "4")
        self.threshold_entry.pack(side="left", padx=5)
        tk.Label(threshold_frame, text="days", font=("Arial", 9)).pack(side="left")

        action_frame = tk.Frame(control_frame)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        self.run_btn = tk.Button(action_frame, text="▶ RUN ANALYSIS", command=self.run_analysis,
                                 bg="#FF9800", fg="white", font=("Arial", 10, "bold"), padx=30, pady=5)
        self.run_btn.pack(side="left", padx=5)
        
        self.pivot_btn = tk.Button(action_frame, text="📊 Show Table", command=self.show_pivot_table,
                                   bg="#ff6a6a", fg="white", font=("Arial", 10,"bold"), padx=20, pady=5, state="disabled")
        self.pivot_btn.pack(side="left", padx=5)
        
        self.save_btn = tk.Button(action_frame, text="💾 Save Results", command=self.save_results,
                                  bg="#F44336", fg="white", font=("Arial", 10,"bold"), padx=20, pady=5, state="disabled")
        self.save_btn.pack(side="left", padx=5)
        
        # Progress bar and status
        progress_frame = tk.Frame(control_frame)
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=300)
        self.progress.pack(side="left", padx=5)
        
        self.status_label = tk.Label(progress_frame, text="Ready", fg="green", font=("Arial", 9))
        self.status_label.pack(side="left", padx=10)
        
        # Preview section
        preview_frame = tk.LabelFrame(main_frame, text="Data Preview", font=("Arial", 10, "bold"))
        preview_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Treeview with scrollbar
        tree_container = tk.Frame(preview_frame)
        tree_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_container)
        scrollbar_y.pack(side="right", fill="y")
        
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        
        self.tree = ttk.Treeview(tree_container, yscrollcommand=scrollbar_y.set, 
                                  xscrollcommand=scrollbar_x.set, height=20)
        self.tree.pack(fill="both", expand=True)
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # Info bar
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill="x")
        
        self.info_label = tk.Label(info_frame, text="", fg="blue", font=("Arial", 9))
        self.info_label.pack(side="left")
    
    def show_date_picker(self, entry_widget):
        """
        Display an interactive calendar picker for date selection.
        
        Opens a popup window with a calendar that allows users to select
        a date and automatically fills the associated entry widget.
        
        Args:
            entry_widget (tk.Entry): The entry widget to populate with the selected date.
        """
        import calendar
        from datetime import datetime
        
        date_win = tk.Toplevel(self.root)
        date_win.title("Select Date")
        date_win.geometry("250x250")
        
        # Get current date from entry or use today
        try:
            current_date = datetime.strptime(entry_widget.get(), "%Y-%m-%d")
        except:
            current_date = datetime.now()
        
        year = current_date.year
        month = current_date.month
        
        def update_calendar():
            # Clear previous calendar
            for widget in cal_frame.winfo_children():
                widget.destroy()
            
            # Month navigation
            month_label.config(text=f"{calendar.month_name[month]} {year}")
            
            # Day buttons
            cal = calendar.monthcalendar(year, month)
            days = ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su']
            for col, day in enumerate(days):
                tk.Label(cal_frame, text=day, font=("Arial", 8, "bold")).grid(row=0, column=col, padx=5, pady=2)
            
            for row, week in enumerate(cal, 1):
                for col, day in enumerate(week):
                    if day == 0:
                        continue
                    btn = tk.Button(cal_frame, text=str(day), width=4, height=1,
                                   command=lambda d=day: select_date(d))
                    btn.grid(row=row, column=col, padx=2, pady=2)
        
        def select_date(day):
            selected = f"{year:04d}-{month:02d}-{day:02d}"
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, selected)
            date_win.destroy()
        
        def prev_month():
            nonlocal year, month
            if month == 1:
                month = 12
                year -= 1
            else:
                month -= 1
            update_calendar()
        
        def next_month():
            nonlocal year, month
            if month == 12:
                month = 1
                year += 1
            else:
                month += 1
            update_calendar()
        
        # Navigation buttons
        nav_frame = tk.Frame(date_win)
        nav_frame.pack(pady=10)
        
        tk.Button(nav_frame, text="◀", command=prev_month).pack(side="left", padx=10)
        month_label = tk.Label(nav_frame, text="", font=("Arial", 10, "bold"))
        month_label.pack(side="left", padx=10)
        tk.Button(nav_frame, text="▶", command=next_month).pack(side="left", padx=10)
        
        # Calendar frame
        cal_frame = tk.Frame(date_win)
        cal_frame.pack(pady=10)
        
        update_calendar()
        
        # Quick buttons
        quick_frame = tk.Frame(date_win)
        quick_frame.pack(pady=10)
        
        def set_today():
            today = datetime.now().strftime("%Y-%m-%d")
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, today)
            date_win.destroy()
        
        tk.Button(quick_frame, text="Today", command=set_today, width=8).pack(side="left", padx=5)
        
    def update_status(self, message, is_error=False):
        """Update status label with progress message"""
        self.status_label.config(text=message, fg="red" if is_error else "green")
        self.root.update_idletasks()
        
    def start_progress(self):
        """Start progress animation"""
        self.progress.start(10)
        self.run_btn.config(state="disabled")
        
    def stop_progress(self):
        """Stop progress animation"""
        self.progress.stop()
        self.run_btn.config(state="normal")
        
    def load_main(self):
        """Load main Excel file"""
        self.file = filedialog.askopenfilename(
            title="Select Main Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if self.file:
            try:
                self.update_status("Loading main file...")
                xls = pd.ExcelFile(self.file)
                self.data = {s: pd.read_excel(xls, s) for s in xls.sheet_names}
                
                # Update status
                sheets_str = ", ".join(list(self.data.keys()))
                self.main_status.config(text=f"✓ {os.path.basename(self.file)}", fg="green")
                self.update_status(f"Loaded sheets: {sheets_str}")
                
                # Build mapping UI
                self.build_mapping_ui()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load main file:\n{str(e)}")
                self.update_status(f"Error: {str(e)}", is_error=True)
                self.file = None
                self.main_status.config(text="No main file loaded", fg="gray")
                
    def load_afst(self):
        """Load distance reference file with mapping"""
        self.afst = filedialog.askopenfilename(
            title="Select Distance Reference File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if self.afst:
            self.dist_status.config(text=f"✓ {os.path.basename(self.afst)}", fg="green")
            self.update_status("Distance file loaded successfully")
            
            # Build mapping UI for afstanden
            self.build_afstanden_mapping_ui()
    
    def build_afstanden_mapping_ui(self):
        """Create mapping window for afstanden file"""
        if not self.afst:
            return
        
        try:
            xls = pd.ExcelFile(self.afst)
            sheets = xls.sheet_names
            
            self.afstanden_mapping_window = tk.Toplevel(self.root)
            self.afstanden_mapping_window.title("Distance File Column Mapping")
            self.afstanden_mapping_window.geometry("600x500")
            
            tk.Label(self.afstanden_mapping_window, 
                    text="Map columns in distance reference file", 
                    font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(self.afstanden_mapping_window, 
                    text="Select the ID column and Distance column for each sheet", 
                    fg="gray").pack(pady=(0, 10))
            
            # Create notebook for multiple sheets
            notebook = ttk.Notebook(self.afstanden_mapping_window)
            notebook.pack(fill="both", expand=True, padx=10, pady=10)
            
            self.afstanden_mapping_vars = {}
            
            for sheet_idx, sheet_name in enumerate(sheets[:2]):  # Only first 2 sheets
                frame = tk.Frame(notebook)
                notebook.add(frame, text=f"Sheet {sheet_idx+1}: {sheet_name}")
                
                # Load the sheet data
                df = pd.read_excel(xls, sheet_name)
                
                # ID column selection
                tk.Label(frame, text="ID Column (e.g., UGent ID, Pers.nr.):", 
                        font=("Arial", 9)).pack(anchor="w", padx=10, pady=5)
                id_var = tk.StringVar()
                id_combo = ttk.Combobox(frame, textvariable=id_var, width=40)
                id_combo["values"] = [""] + list(df.columns)
                id_combo.pack(padx=10, pady=5)
                self.afstanden_mapping_vars[f"sheet{sheet_idx+1}_id"] = id_var
                
                # Distance column selection
                tk.Label(frame, text="Distance Column (in km):", 
                        font=("Arial", 9)).pack(anchor="w", padx=10, pady=5)
                dist_var = tk.StringVar()
                dist_combo = ttk.Combobox(frame, textvariable=dist_var, width=40)
                dist_combo["values"] = [""] + list(df.columns)
                dist_combo.pack(padx=10, pady=5)
                self.afstanden_mapping_vars[f"sheet{sheet_idx+1}_distance"] = dist_var
                
                # Show preview
                tk.Label(frame, text="Preview (first 5 rows):", 
                        font=("Arial", 9, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
                preview_text = tk.Text(frame, height=5, width=50)
                preview_text.pack(padx=10, pady=5)
                preview_text.insert("1.0", df.head().to_string())
                preview_text.config(state="disabled")
            
            # Buttons
            btn_frame = tk.Frame(self.afstanden_mapping_window)
            btn_frame.pack(pady=10)
            
            tk.Button(btn_frame, text="Save Mapping", command=self.save_afstanden_mapping,
                     bg="#4CAF50", fg="white", padx=20, pady=5).pack(side="left", padx=5)
            tk.Button(btn_frame, text="Skip Mapping", command=self.afstanden_mapping_window.destroy,
                     bg="#9E9E9E", fg="white", padx=20, pady=5).pack(side="left", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load distance file for mapping:\n{str(e)}")
    
    def save_afstanden_mapping(self):
        """Save afstanden column mapping"""
        self.afstanden_mapping = {
            "sheet1": {
                "ID": self.afstanden_mapping_vars.get("sheet1_id", tk.StringVar()).get(),
                "Distance": self.afstanden_mapping_vars.get("sheet1_distance", tk.StringVar()).get()
            },
            "sheet2": {
                "ID": self.afstanden_mapping_vars.get("sheet2_id", tk.StringVar()).get(),
                "Distance": self.afstanden_mapping_vars.get("sheet2_distance", tk.StringVar()).get()
            }
        }
        
        messagebox.showinfo("Success", "Distance file mapping saved!")
        self.afstanden_mapping_window.destroy()
            
    def build_mapping_ui(self):
        """Create column mapping window"""
        self.map_window = tk.Toplevel(self.root)
        self.map_window.title("Column Mapping - Required Fields")
        self.map_window.geometry("700x700")
        
        # Instructions
        tk.Label(self.map_window, text="Map your columns to required fields:", 
                font=("Arial", 12, "bold")).pack(pady=10)
        tk.Label(self.map_window, text="Select the corresponding column from your Excel sheets", 
                fg="gray").pack(pady=(0, 10))
        
        # Canvas with scrollbar for mapping fields
        canvas = tk.Canvas(self.map_window)
        scrollbar = ttk.Scrollbar(self.map_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Store mapping variables
        self.mapping_vars = {}
        
        row = 0
        for sheet, df in self.data.items():
            if sheet not in REQUIRED_FIELDS:
                continue
                
            # Sheet header
            tk.Label(scrollable_frame, text=f"📋 {sheet}", 
                    font=("Arial", 10, "bold"), fg="#2196F3").grid(row=row, column=0, columnspan=3, 
                                                                   sticky="w", pady=(10, 5))
            row += 1
            
            for field in REQUIRED_FIELDS[sheet]:
                # Field label with indicator for critical fields
                if field in ["Pers.nr.", "UGent ID"]:
                    field_text = f"🔑 {field} *"
                    field_color = "red"
                elif field == "Plaats":
                    field_text = f"📍 {field}"
                    field_color = "darkgreen"
                else:
                    field_text = f"  • {field}"
                    field_color = "black"
                
                tk.Label(scrollable_frame, text=field_text, 
                        font=("Arial", 9), fg=field_color).grid(row=row, column=0, sticky="w", padx=20, pady=2)
                
                # Add hint for specific fields
                hint_text = ""
                if field == "Plaats":
                    hint_text = "(woonplaats/location for distance mapping)"
                elif field == "Pers.nr.":
                    hint_text = "(personal number - used for ID linking)"
                elif field == "UGent ID":
                    hint_text = "(university ID - will be linked)"
                elif field == "E-mail":
                    hint_text = "(for email column in results)"
                
                if hint_text:
                    hint_label = tk.Label(scrollable_frame, text=hint_text, 
                                          font=("Arial", 7), fg="gray")
                    hint_label.grid(row=row, column=2, sticky="w", padx=5)
                
                # Dropdown for column selection
                var = tk.StringVar()
                combo = ttk.Combobox(scrollable_frame, textvariable=var, width=35)
                combo["values"] = [""] + list(df.columns)
                combo.grid(row=row, column=1, padx=10, pady=2, sticky="w")
                
                self.mapping_vars[(sheet, field)] = var
                row += 1
        
        # Add validation info
        info_frame = tk.Frame(scrollable_frame)
        info_frame.grid(row=row, column=0, columnspan=3, pady=20)
        tk.Label(info_frame, text="⚠️  Important: 'Pers.nr.' and 'UGent ID' fields must be mapped correctly", 
                font=("Arial", 9), fg="red").pack()
        tk.Label(info_frame, text="These are used to link data across all sheets", 
                font=("Arial", 8), fg="gray").pack()
        
        row += 1
        
        # Buttons
        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        tk.Button(btn_frame, text="✓ Save Mapping", command=self.save_mapping,
                 bg="#4CAF50", fg="white", padx=20, pady=5).pack(side="left", padx=5)
        tk.Button(btn_frame, text="✗ Skip Mapping (Use Original Names)", command=self.map_window.destroy,
                 bg="#FF9800", fg="white", padx=20, pady=5).pack(side="left", padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def save_mapping(self):
        """Save column mapping"""
        self.mapping = {}
        
        for (sheet, field), var in self.mapping_vars.items():
            if sheet not in self.mapping:
                self.mapping[sheet] = {}
            self.mapping[sheet][field] = var.get()
        
        messagebox.showinfo("Success", f"Mapping saved for {len(self.mapping)} sheets!")
        self.map_window.destroy()
        
    def run_analysis(self):
        """
        Execute the full mobility analysis pipeline in a background thread.
        
        Validates input files and dates, loads the engine, applies column mapping,
        processes the data, and updates the UI with results.
        """
        if not self.file or not self.afst:
            messagebox.showerror("Error", "Please load both Main Excel and Distance files first!")
            return
        
        # Validate date format
        try:
            start_date = pd.to_datetime(self.start_entry.get())
            end_date = pd.to_datetime(self.end_entry.get())
            
            if start_date > end_date:
                messagebox.showerror("Error", "Start date cannot be after end date!")
                return
                
        except Exception as e:
            messagebox.showerror("Error", 
                                f"Invalid date format!\n\n"
                                f"Please use YYYY-MM-DD format.\n"
                                f"Example: 2024-10-01\n\n"
                                f"Error: {str(e)}")
            return
        
        def run_thread():
            try:
                self.start_progress()
                self.update_status("Initializing analysis engine...")
                
                # Create engine
                self.engine = MobilityEngine(self.file, self.afst, 
                                            self.start_entry.get(), self.end_entry.get(),
                                            bike_threshold=int(self.threshold_entry.get() or 4))
                
                self.update_status("Loading data...")
                self.engine.load_data()
                
                # Check if MOB0 exists
                if "MOB0" not in self.engine.data:
                    raise Exception("MOB0 sheet not found in main file! This sheet is required for linking IDs.")
                
                # Show available sheets
                available_sheets = list(self.engine.data.keys())
                self.update_status(f"Found sheets: {', '.join(available_sheets)}")
                
                self.update_status("Applying column mapping...")
                self.engine.data = apply_mapping(self.engine.data, self.mapping)
                
                self.update_status("Loading distance files with mapping...")
                if hasattr(self, 'afstanden_mapping') and self.afstanden_mapping:
                    self.engine.load_afstanden_with_mapping(
                        self.afstanden_mapping.get("sheet1", {}),
                        self.afstanden_mapping.get("sheet2", {})
                    )
                else:
                    self.engine.load_afstanden_with_mapping()
                
                self.update_status("Running full analysis pipeline (standardize → attach → deduplicate)...")
                self.engine.run()
                
                self.update_status("Building final conclusions...")
                self.engine.build_conclusie()
                
                self.df = self.engine.conclusie
                
                # Generate pivot table
                self.pivot_table = self.engine.get_pivot_table()
                
                if self.df is None or len(self.df) == 0:
                    raise Exception("Analysis produced no results. Check if date range is valid.")
                
                self.update_status(f"✓ Analysis complete! Found {len(self.df)} records")
                
                # Update UI in main thread
                self.root.after(0, self.analysis_complete)
                
            except Exception as e:
                error_msg = str(e)
                error_trace = traceback.format_exc()
                
                print("=" * 50)
                print("ERROR DETAILS:")
                print(error_trace)
                print("=" * 50)
                
                detailed_error = f"{error_msg}\n\nCheck console for full details."
                self.root.after(0, lambda msg=detailed_error: self.analysis_error(msg))
        
        # Start the thread
        threading.Thread(target=run_thread, daemon=True).start()
        
    def analysis_complete(self):
        """Handle successful analysis completion"""
        self.stop_progress()
        self.update_status(f"✅ Analysis complete! Found {len(self.df)} records")
        
        # Show preview
        self.show_preview(self.df)
        
        # Enable buttons
        self.pivot_btn.config(state="normal")
        self.save_btn.config(state="normal")
        
        # Show info
        summary = self.get_summary_stats()
        self.info_label.config(text=summary)
        
        messagebox.showinfo("Success", 
                           f"Analysis completed successfully!\n\n"
                           f"Total employees: {len(self.df)}\n"
                           f"Transport modes analyzed\n"
                           f"Results ready for viewing and export")
        
    def analysis_error(self, error_msg):
        """Handle analysis error"""
        self.stop_progress()
        self.update_status(f"❌ Error: {error_msg}", is_error=True)
        messagebox.showerror("Analysis Error", f"Failed to run analysis:\n\n{error_msg}")
        
    def show_preview(self, df):
        """
        Display a preview of the analysis results in the treeview widget.
        
        Shows the first 500 rows of the dataframe with sortable columns.
        
        Args:
            df (pd.DataFrame): The dataframe to preview.
        """
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        for col in columns:
            width = min(150, max(80, len(str(col)) * 10))
            self.tree.column(col, width=width, minwidth=50)
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c, False))
        
        for _, row in df.head(500).iterrows():
            values = [str(v)[:50] + "..." if len(str(v)) > 50 else str(v) for v in row]
            self.tree.insert("", "end", values=values)
        
    def sort_column(self, col, reverse):
        """
        Sort the treeview by a selected column.
        
        Args:
            col (str): The column name to sort by.
            reverse (bool): If True, sort in descending order; otherwise ascending.
        """
        data_list = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        data_list.sort(reverse=reverse)
        
        for index, (val, child) in enumerate(data_list):
            self.tree.move(child, '', index)
        
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))
        
    def get_summary_stats(self):
        """
        Generate summary statistics from the analysis results.
        
        Returns:
            str: A formatted string with key statistics.
        """
        if self.df is None or len(self.df) == 0:
            return ""
        
        stats = []
        stats.append(f"📊 Total: {len(self.df)} employees")
        
        if "vervoerswijze" in self.df.columns:
            modes = self.df["vervoerswijze"].value_counts()
            top_mode = modes.index[0] if len(modes) > 0 else "N/A"
            stats.append(f"🚲 Most common: {top_mode} ({modes[top_mode]} users)")
        
        return " | ".join(stats)
    
    def show_pivot_table(self):
        """
        Display a pivot table showing transport mode distribution by site.
        
        Opens a new window with a sortable table and export functionality.
        """
        if self.pivot_table is None or len(self.pivot_table) == 0:
            messagebox.showwarning("No Data", "No pivot table data available. Please run analysis first!")
            return
        
        win = tk.Toplevel(self.root)
        win.title("Transport Modes by Site - Pivot Table")
        win.geometry("1000x600")
        
        # Add title
        title_frame = tk.Frame(win)
        title_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Label(title_frame, text="Transport Modes Distribution by Site", 
                 font=("Arial", 14, "bold")).pack()
        tk.Label(title_frame, text=f"Analysis Period: {self.start_entry.get()} to {self.end_entry.get()}", 
                 font=("Arial", 9), fg="gray").pack()
        
        # Create frame for table
        table_frame = tk.Frame(win)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create treeview with scrollbars
        tree_container = tk.Frame(table_frame)
        tree_container.pack(fill="both", expand=True)
        
        scrollbar_y = ttk.Scrollbar(tree_container)
        scrollbar_y.pack(side="right", fill="y")
        
        scrollbar_x = ttk.Scrollbar(tree_container, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        
        # Create treeview
        pivot_tree = ttk.Treeview(tree_container, yscrollcommand=scrollbar_y.set, 
                                   xscrollcommand=scrollbar_x.set)
        pivot_tree.pack(fill="both", expand=True)
        
        scrollbar_y.config(command=pivot_tree.yview)
        scrollbar_x.config(command=pivot_tree.xview)
        
        # Configure columns
        columns = list(self.pivot_table.columns)
        pivot_tree["columns"] = columns
        pivot_tree["show"] = "headings"
        
        # Add site column as first column
        pivot_tree.column("#0", width=150, minwidth=100)
        pivot_tree.heading("#0", text="Site")
        
        # Add transport mode columns
        for col in columns:
            pivot_tree.column(col, width=100, minwidth=80)
            pivot_tree.heading(col, text=col)
        
        # Add data
        for idx, row in self.pivot_table.iterrows():
            values = [row[col] for col in columns]
            pivot_tree.insert("", "end", text=idx, values=values)
        
        # Add export button
        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)
        
        def export_pivot():
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile="pivot_table_transport_modes.xlsx"
            )
            if path:
                self.pivot_table.to_excel(path)
                messagebox.showinfo("Success", f"Pivot table saved to:\n{path}")
        
        tk.Button(btn_frame, text="Export to Excel", command=export_pivot,
                 bg="#4CAF50", fg="white", padx=20, pady=5).pack()
        
    def save_results(self):
        """
        Export analysis results to an Excel workbook.
        
        Creates a multi-sheet Excel file containing the main results,
        pivot table, summary statistics, and mode breakdown.
        """
        if self.df is None:
            messagebox.showwarning("No Data", "No results to save. Please run analysis first!")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"mobility_results_{timestamp}.xlsx"
        
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_name
        )
        
        if path:
            try:
                self.update_status("Saving results...")
                
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    self.df.to_excel(writer, sheet_name='Mobility Results', index=False)
                    
                    if self.pivot_table is not None and len(self.pivot_table) > 0:
                        self.pivot_table.to_excel(writer, sheet_name='Pivot Table - Transport by Site')
                    
                    summary = pd.DataFrame({
                        'Metric': ['Total Employees', 'Analysis Period', 'Average Distance', 'Unique Sites'],
                        'Value': [
                            len(self.df),
                            f"{self.start_entry.get()} to {self.end_entry.get()}",
                            f"{self.df['afstand'].mean():.1f} km" if 'afstand' in self.df.columns else 'N/A',
                            len(self.df['site'].unique()) if 'site' in self.df.columns else 'N/A'
                        ]
                    })
                    summary.to_excel(writer, sheet_name='Summary', index=False)
                    
                    if 'vervoerswijze' in self.df.columns:
                        mode_summary = self.df['vervoerswijze'].value_counts().reset_index()
                        mode_summary.columns = ['Transport Mode', 'Count']
                        mode_summary['Percentage'] = (mode_summary['Count'] / len(self.df) * 100).round(1)
                        mode_summary.to_excel(writer, sheet_name='Mode Breakdown', index=False)
                
                self.update_status(f"✅ Results saved to {os.path.basename(path)}")
                messagebox.showinfo("Success", f"Results saved successfully!\n\nLocation: {path}")
                
            except Exception as e:
                self.update_status(f"Error saving: {str(e)}", is_error=True)
                messagebox.showerror("Save Error", f"Failed to save results:\n{str(e)}")


if __name__ == "__main__":
    """
    Application entry point.
    
    Creates the root Tkinter window, initializes the App, and starts
    the main event loop. Catches any critical errors and logs them.
    """
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except Exception as e:
        error_text = traceback.format_exc()
        log_path = os.path.join(get_app_path(), "error.log")
        
        with open(log_path, "w") as f:
            f.write(f"Error: {str(e)}\n\n") 
            f.write(error_text)
        
        messagebox.showerror("Critical Error", f"Something went wrong.\nCheck:\n{log_path}")