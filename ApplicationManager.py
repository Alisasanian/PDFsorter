import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, font
import ctypes
import os
import sys
import subprocess
import threading
import shutil
import re
from pathlib import Path
from datetime import datetime

dirpath = Path(__file__).parent.as_posix()

class PDFSorterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Sorting Tool")
        self.root.geometry("800x700")
        
        # Set modern colors
        self.bg_color = "#f0f0f0"
        self.primary_color = "#2c3e50"
        self.accent_color = "#3498db"
        self.success_color = "#27ae60"
        self.error_color = "#e74c3c"
        
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.selected_directory = tk.StringVar()
        self.input_method = tk.StringVar(value="airtable")
        self.airtable_link = tk.StringVar()
        self.csv_file_path = tk.StringVar()
        self.is_processing = False
        self.current_progress = 0
        self.terminal_visible = tk.BooleanVar(value=False)
        self.output_visible = tk.BooleanVar(value=False)
        self.processing_errors = []
        self.capture_summary = False
        self.summary_lines = []
        
        # Get the directory where this script is located
        self.script_dir = Path(__file__).parent
        
        # Create necessary directories
        self.create_directories()
        
        # Create main UI
        self.create_widgets()
    
    def create_directories(self):
        """Create necessary directories if they don't exist"""
        directories = [
            self.script_dir / "ReferenceCSV",
            self.script_dir / "PDFsToProcess",
            self.script_dir / "_workingdata_",
            self.script_dir / "_workingdata_" / "_bloatedcache_",
            self.script_dir / "_workingdata_" / "_pdfcache_",
            self.script_dir / "_workingdata_" / "_indexdataset_",
            self.script_dir / "_workingdata_" / "_indexdataset_" / "combined_data",
            self.script_dir / "SortedPDFs"
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
        
    def create_widgets(self):
        """Create all GUI widgets with modern styling"""
        
        # Main container with padding
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(
            main_frame, 
            text="PDF Sorting Tool",
            font=("Arial", 36, "bold"),
            bg=self.bg_color,
            fg=self.primary_color
        )
        title_label.pack(pady=(0, 20))
        
        # Step 1: Directory Selection
        step1_frame = self.create_section_frame(main_frame, "Select PDF Directory")
        
        dir_frame = tk.Frame(step1_frame, bg="white")
        dir_frame.pack(fill="x", pady=(10, 0))
        
        self.dir_entry = tk.Entry(
            dir_frame,
            textvariable=self.selected_directory,
            font=("Arial", 11),
            state="readonly",
            bg="#f8f8f8"
        )
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        
        self.browse_btn = self.create_button(
            dir_frame,
            "Browse",
            self.browse_directory,
            width=15
        )
        self.browse_btn.pack(side="right", padx=(5, 10), pady=10)
        
        # Step 2: Master Sheet Selection
        step2_frame = self.create_section_frame(main_frame, "Select Master Sheet Source")
        
        # Radio buttons for input method
        radio_frame = tk.Frame(step2_frame, bg="white")
        radio_frame.pack(fill="x", pady=(10, 0))
        
        airtable_radio = tk.Radiobutton(
            radio_frame,
            text="Airtable Link",
            variable=self.input_method,
            value="airtable",
            font=("Arial", 11),
            bg="white",
            activebackground="white",
            command=self.toggle_input_method
        )
        airtable_radio.pack(side="left", padx=(10, 20), pady=10)
        
        csv_radio = tk.Radiobutton(
            radio_frame,
            text="CSV File",
            variable=self.input_method,
            value="csv",
            font=("Arial", 11),
            bg="white",
            activebackground="white",
            command=self.toggle_input_method
        )
        csv_radio.pack(side="left", pady=10)
        
        # Airtable input
        self.airtable_frame = tk.Frame(step2_frame, bg="white")
        self.airtable_frame.pack(fill="x", pady=(5, 10))
        
        self.airtable_entry = tk.Entry(
            self.airtable_frame,
            textvariable=self.airtable_link,
            font=("Arial", 11),
            bg="#f8f8f8"
        )
        self.airtable_entry.pack(fill="x", padx=10, pady=10)
        
        # CSV input
        self.csv_frame = tk.Frame(step2_frame, bg="white")
        
        csv_input_frame = tk.Frame(self.csv_frame, bg="white")
        csv_input_frame.pack(fill="x")
        
        self.csv_entry = tk.Entry(
            csv_input_frame,
            textvariable=self.csv_file_path,
            font=("Arial", 11),
            state="readonly",
            bg="#f8f8f8"
        )
        self.csv_entry.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=10)
        
        self.csv_browse_btn = self.create_button(
            csv_input_frame,
            "Browse CSV",
            self.browse_csv,
            width=15
        )
        self.csv_browse_btn.pack(side="right", padx=(5, 10), pady=10)
        
        # Initially show only Airtable input
        self.toggle_input_method()
        
        # Sort button
        self.sort_btn = self.create_button(
            main_frame,
            "SORT",
            self.start_sorting,
            width=10,
            height=1,
            font_size=28,
            bg_color=self.success_color
        )
        self.sort_btn.pack(pady=20)
        
        # Progress section with toggle
        progress_frame = self.create_section_frame(main_frame, "Progress")
        
        # Progress bar container with toggle button
        progress_container = tk.Frame(progress_frame, bg="white")
        progress_container.pack(fill="x", padx=10, pady=(10, 5))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_container,
            variable=self.progress_var,
            maximum=100,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill="x", padx=(0, 1), pady=(0, 1))
        
        # Toggle button at bottom right of progress bar
        self.toggle_btn = tk.Button(
            progress_container,
            text="open log ▼",
            font=("Arial", 10),
            bg="white",
            fg=self.primary_color,
            activebackground="white",
            activeforeground=self.accent_color,
            relief="flat",
            bd=0,
            cursor="hand2",
            command=self.toggle_terminal
        )
        self.toggle_btn.pack(side="right")
        
        # Add hover effect to toggle button
        self.toggle_btn.bind("<Enter>", lambda e: self.toggle_btn.config(fg=self.accent_color))
        self.toggle_btn.bind("<Leave>", lambda e: self.toggle_btn.config(fg=self.primary_color))
        
        # Progress label
        self.progress_label = tk.Label(
            progress_frame,
            text="Ready to process",
            font=("Arial", 14),
            bg="white",
            fg=self.primary_color
        )
        self.progress_label.pack(pady=(0, 10))
        
        # Terminal Output section (collapsible)
        self.terminal_frame = tk.Frame(main_frame, bg="white", relief="flat", borderwidth=1)
        self.terminal_frame.configure(highlightbackground="#e0e0e0", highlightthickness=1)
        
        # Terminal title
        terminal_title_label = tk.Label(
            self.terminal_frame,
            text="Terminal Output",
            font=("Arial", 12, "bold"),
            bg="white",
            fg=self.primary_color
        )
        terminal_title_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Terminal text area
        self.terminal_text = scrolledtext.ScrolledText(
            self.terminal_frame,
            height=12,
            font=("Consolas", 12),
            bg="#1e1e1e",
            fg="#00ff00",
            insertbackground="#00ff00",
            wrap=tk.WORD,
            state="disabled"  # Make text read-only
        )
        self.terminal_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Create output frame before toggling terminal
        # Error/Output display
        self.output_frame = self.create_section_frame(main_frame, "Log")
        
        self.output_text = scrolledtext.ScrolledText(
            self.output_frame,
            height=6,
            font=("Consolas", 12),
            bg="#f8f8f8",
            wrap=tk.WORD
        )
        self.output_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Now hide terminal after output_text is created
        self.toggle_terminal()
        
        # Configure custom progressbar style
        style = ttk.Style()
        style.configure(
            "Custom.Horizontal.TProgressbar",
            background=self.accent_color,
            troughcolor="#e0e0e0",
            bordercolor="#d0d0d0",
            lightcolor=self.accent_color,
            darkcolor=self.accent_color
        )
        
    def create_section_frame(self, parent, title):
        """Create a styled section frame with title"""
        frame = tk.Frame(parent, bg="white", relief="flat", borderwidth=1)
        frame.pack(fill="x", pady=(0, 15))
        
        # Add subtle border effect
        frame.configure(highlightbackground="#e0e0e0", highlightthickness=1)
        
        # Section title
        title_label = tk.Label(
            frame,
            text=title,
            font=("Arial", 12, "bold"),
            bg="white",
            fg=self.primary_color
        )
        title_label.pack(anchor="w", padx=10, pady=(10, 0))
        
        return frame
        
    def create_button(self, parent, text, command, width=None, height=1, font_size=11, bg_color=None):
        """Create a modern styled button"""
        if bg_color is None:
            bg_color = self.accent_color
            
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Arial", font_size),
            bg=bg_color,
            fg="white",
            activebackground=self.primary_color,
            activeforeground="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            width=width,
            height=height
        )
        
        # Add hover effect
        btn.bind("<Enter>", lambda e: btn.config(bg=self.primary_color))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))
        
        return btn
    
    def toggle_terminal(self):
        """Toggle the visibility of the terminal output"""
        if self.terminal_visible.get():
            # Hide terminal
            self.terminal_frame.pack_forget()
            self.terminal_visible.set(False)
            self.output_frame.pack(fill="both", expand=True, pady=(0, 5))
            self.output_visible.set(True)
            self.toggle_btn.config(text="close log ▲")
            self.root.geometry("800x900")
        else:
            # Show terminal
            self.terminal_frame.pack(fill="both", expand=True, before=self.output_text.master, pady=(0, 5))
            self.terminal_visible.set(True)
            self.output_frame.pack_forget()
            self.output_visible.set(False)
            self.toggle_btn.config(text="open log ▼")
            self.root.geometry("800x700")
        
    def browse_directory(self):
        """Open directory browser dialog"""
        global directory
        directory = filedialog.askdirectory(title="Select Directory Containing PDFs")
        if directory:
            self.selected_directory.set(directory)
            self.log_message(f"Selected directory: {directory}", "info")
            
    def browse_csv(self):
        """Open file browser for CSV selection"""
        filename = filedialog.askopenfilename(
            title="Select Master CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_file_path.set(filename)
            self.log_message(f"Selected CSV: {filename}", "info")
            
    def toggle_input_method(self):
        """Toggle between Airtable and CSV input methods"""
        if self.input_method.get() == "airtable":
            self.airtable_frame.pack(fill="x", pady=(5, 10))
            self.csv_frame.pack_forget()
        else:
            self.csv_frame.pack(fill="x", pady=(5, 10))
            self.airtable_frame.pack_forget()
            
    def log_message(self, message, msg_type="info"):
        """Log a message to the output text area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Configure text tags for different message types
        self.output_text.tag_config("info", foreground=self.primary_color)
        self.output_text.tag_config("success", foreground=self.success_color)
        self.output_text.tag_config("error", foreground=self.error_color)
        self.output_text.tag_config("warning", foreground="#f39c12")
        
        # Insert message with appropriate tag
        self.output_text.insert(tk.END, f"[{timestamp}] {message}\n", msg_type)
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def terminal_output(self, message):
        """Output message to terminal window"""
        self.terminal_text.config(state="normal")  # Enable for writing
        self.terminal_text.insert(tk.END, message)
        self.terminal_text.see(tk.END)
        self.terminal_text.config(state="disabled")  # Disable after writing
        self.root.update_idletasks()

    def show_missing_drawings_warning(self, total_drawings, found_drawings, missing_drawings, missing_list):
        """Show a warning window for missing drawings"""
        warning_window = tk.Toplevel(self.root)
        warning_window.title("SORTED PDF WARNING")
        warning_window.geometry("600x500")
        warning_window.configure(bg="white")
        
        # Warning icon and title
        title_frame = tk.Frame(warning_window, bg="white")
        title_frame.pack(pady=20)
        
        warning_label = tk.Label(
            title_frame,
            text="SORTING SUMMARY",
            font=("Arial", 24, "bold"),
            bg="white",
            fg=self.error_color
        )
        warning_label.pack()
        
        # Summary info
        info_frame = tk.Frame(warning_window, bg="white")
        info_frame.pack(pady=5)
        
        summary_text = f"Total drawings from master CSV: {total_drawings}\n"
        summary_text += f"Found in index CSV: {found_drawings}\n"
        summary_text += f"Missing from CSV: {missing_drawings}"
        
        summary_label = tk.Label(
            info_frame,
            text=summary_text,
            font=("Arial", 16),
            bg="white",
            fg=self.primary_color
        )
        summary_label.pack()
        
        # Missing drawings list
        if missing_list:
            list_frame = tk.Frame(warning_window, bg="white")
            list_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
            list_label = tk.Label(
                list_frame,
                text="Missing drawings:",
                font=("Arial", 16, "bold"),
                bg="white",
                fg=self.error_color
            )
            list_label.pack(anchor="w")
            
            # Scrollable text for missing drawings
            missing_text = scrolledtext.ScrolledText(
                list_frame,
                height=10,
                font=("Consolas", 16),
                bg="#f8f8f8",
                wrap=tk.WORD,
                state="disabled"
            )
            missing_text.pack(fill="both", expand=True, pady=(5, 0))
            
            missing_text.config(state="normal")
            missing_text.insert(tk.END, ", ".join(missing_list))
            missing_text.config(state="disabled")
        
        # OK button
        ok_button = tk.Button(
            warning_window,
            text="OK",
            command=warning_window.destroy,
            font=("Arial", 12),
            bg=self.accent_color,
            fg="white",
            relief="flat",
            bd=0,
            cursor="hand2",
            width=10
        )
        ok_button.pack(pady=20)
        
        # Center the window
        warning_window.transient(self.root)
        warning_window.grab_set()
        
    def validate_inputs(self):
        """Validate user inputs before processing"""
        if not self.selected_directory.get():
            messagebox.showerror("Error", "Please select a directory containing PDFs")
            return False
            
        if self.input_method.get() == "airtable":
            if not self.airtable_link.get():
                messagebox.showerror("Error", "Please enter an Airtable link")
                return False
        else:
            if not self.csv_file_path.get():
                messagebox.showerror("Error", "Please select a CSV file")
                return False
                
        return True
        
    def start_sorting(self):
        """Start the PDF sorting process"""
        if not self.validate_inputs():
            return
            
        if self.is_processing:
            messagebox.showwarning("Warning", "Processing is already in progress")
            return
            
        # Reset error tracking
        self.processing_errors = []
        
        # Disable controls during processing
        self.is_processing = True
        self.sort_btn.config(state="disabled", text="Processing...")
        self.browse_btn.config(state="disabled")
        self.csv_browse_btn.config(state="disabled")
        
        # Clear outputs
        self.output_text.delete(1.0, tk.END)
        self.terminal_text.delete(1.0, tk.END)
        
        # Show terminal if hidden
        if not self.terminal_visible.get():
            self.toggle_terminal()
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()
        
    def process_pdfs(self):
        """Main processing function that runs the Python scripts in sequence"""
        try:
            # Initial terminal output
            self.terminal_output("Sorting PDFs. . .\n")
            
            # Copy PDFs to processing directory
            self.prepare_pdfs()
            
            # Prepare reference CSV if needed
            if self.input_method.get() == "csv":
                self.prepare_csv()
            
            # List of scripts to run with their progress ranges
            scripts = [
                ("PDFcombiner.py", 0, 20, "Combining PDFs..."),
                ("ExpandedPDFdrawingNumberCrop.py", 20, 40, "Cropping drawing numbers..."),
                ("cropToJPEGcachePDF.py", 40, 60, "Converting to JPEG cache..."),
                ("CacheOCR.py", 60, 80, "Running OCR..."),
                (self.get_sorter_script(), 80, 100, "Sorting PDFs...")
            ]
            
            # Track overall success
            all_successful = True
            
            # Run each script
            for script, start_progress, end_progress, status_msg in scripts:
                self.update_progress(start_progress, status_msg)
                
                # Add script separator to terminal
                separator = "=" * 86
                self.terminal_output(f"\n{separator}\n")
                self.terminal_output(f"{script}\n")
                self.terminal_output(f"{separator}\n\n")
                
                success = self.run_script(script, start_progress, end_progress)
                
                if not success:
                    all_successful = False
                    self.log_message(f"Error running {script}", "error")
                    self.processing_errors.append(f"Failed to run {script}")
                    break
            
            # Check for errors collected during processing
            if self.processing_errors or not all_successful:
                self.update_progress(self.progress_var.get(), "Error in sorting PDF :(\n[check logs for error codes]")
                self.log_message("PDF sorting failed due to errors", "error")
                self.terminal_output("\n\nProcessing failed with errors!\n")
                self.toggle_terminal()
                
                # Show summary of errors
                for error in self.processing_errors:
                    self.terminal_output(f"ERROR: {error}\n")
            else:
                self.update_progress(100, "Processing complete!")
                self.log_message("PDF sorting completed successfully!", "success")
                self.terminal_output("\n\nProcessing complete!\n")
            
            # Check if we have missing drawings from the last script
            if hasattr(self, 'summary_lines') and self.summary_lines:
                # Parse summary lines
                total_drawings = 0
                found_drawings = 0
                missing_count = 0
                missing_list = []
                
                for line in self.summary_lines:
                    if "Total drawings from master CSV:" in line:
                        total_drawings = int(re.search(r':\s*(\d+)', line).group(1))
                    elif "Found in index CSV:" in line:
                        found_drawings = int(re.search(r':\s*(\d+)', line).group(1))
                    elif "Missing from CSV:" in line:
                        missing_count = int(re.search(r':\s*(\d+)', line).group(1))
                    elif "Missing drawings:" in line:
                        # Extract the list of missing drawings
                        missing_str = line.split("Missing drawings:")[1].strip()
                        missing_list = [d.strip() for d in missing_str.split(',')]
                
                # Show warning if there are missing drawings
                if missing_count > 0:
                    self.root.after(100, lambda: self.show_missing_drawings_warning(
                        total_drawings, found_drawings, missing_count, missing_list
                    ))

            Sorted_PDF = [f for f in Path(f"{dirpath}/SortedPDFs/").iterdir() if f.suffix.lower() == '.pdf']
            shutil.move(Path(Sorted_PDF[0]), Path(f"{directory}/SORTED_combined.pdf"))
            shutil.rmtree(f"{dirpath}/SortedPDFs/")
            
        except Exception as e:
            self.log_message(f"Processing error: {str(e)}", "error")
            self.terminal_output(f"\n\nERROR: {str(e)}\n")
            self.update_progress(self.progress_var.get(), "Error in sorting PDF :(\n[check logs for error codes]")
            self.toggle_terminal()
            
        finally:
            # Re-enable controls
            self.root.after(0, self.reset_controls)
            
    def prepare_pdfs(self):
        """Copy PDFs from selected directory to PDFsToProcess directory"""
        source_dir = Path(self.selected_directory.get())
        dest_dir = self.script_dir / "PDFsToProcess"
        
        # Create directory if it doesn't exist
        dest_dir.mkdir(parents=True, exist_ok=True)
        
        # Clear existing files
        for file in dest_dir.glob("*.pdf"):
            file.unlink()
            
        # Copy PDF files
        pdf_count = 0
        for pdf_file in source_dir.glob("*.pdf"):
            shutil.copy2(pdf_file, dest_dir)
            pdf_count += 1
            
        self.log_message(f"Copied {pdf_count} PDF(s) to processing directory", "info")
        self.terminal_output(f"Copied {pdf_count} PDF(s) to processing directory\n")
        
    def prepare_csv(self):
        """Copy selected CSV to ReferenceCSV directory"""
        source_csv = Path(self.csv_file_path.get())
        dest_dir = self.script_dir / "ReferenceCSV"
        
        # Create directory if it doesn't exist
        dest_dir.mkdir(parents=True, exist_ok=True)
        
        # Clear existing files
        for file in dest_dir.glob("*.csv"):
            file.unlink()
            
        # Copy CSV file
        shutil.copy2(source_csv, dest_dir / "reference.csv")
        self.log_message("Copied reference CSV file", "info")
        self.terminal_output("Copied reference CSV file\n")
        
    def get_sorter_script(self):
        """Determine which sorter script to use based on input method"""
        if self.input_method.get() == "airtable":
            # Need to update the Airtable URL in PDFpageSorter.py
            self.update_airtable_url()
            return "PDFpageSorter.py"
        else:
            return "PDFpageSortercsv.py"
            
    def update_airtable_url(self):
        """Update the Airtable URL in PDFpageSorter.py"""
        script_path = self.script_dir / "PDFpageSorter.py"
        
        # Read the script content
        with open(script_path, 'r') as f:
            content = f.read()
            
        # Replace the tableurl line
        new_url = self.airtable_link.get()
        content = re.sub(
            r'tableurl = ".*?"',
            f'tableurl = "{new_url}"',
            content
        )
        
        # Write back
        with open(script_path, 'w') as f:
            f.write(content)
            
        self.log_message("Updated Airtable URL", "info")
        self.terminal_output("Updated Airtable URL\n")
        
    def run_script(self, script_name, start_progress, end_progress):
        """Run a Python script and capture its output"""
        try:
            script_path = self.script_dir / script_name
            self.log_message(f"Running {script_name}...", "info")
            
            # Run the script with UTF-8 encoding
            process = subprocess.Popen(
                [sys.executable, str(script_path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                cwd=str(self.script_dir),
                encoding='utf-8',
                errors='replace'  # Replace invalid characters instead of failing
            )
            
            # Track if this script has errors
            script_has_errors = False
            
            # Read output line by line
            for line in process.stdout:
                if line:
                    # Output to terminal
                    self.terminal_output(line)
                    
                    line_stripped = line.strip()
                    if line_stripped:
                        # Check for progress indicators in the output
                        self.parse_script_output(line_stripped, start_progress, end_progress)
                        
                        # Check for sorting summary (ADD THIS AFTER line_stripped is defined)
                        if "--- SORTING SUMMARY ---" in line_stripped:
                            self.capture_summary = True
                            self.summary_lines = []
                        elif hasattr(self, 'capture_summary') and self.capture_summary and line_stripped:
                            self.summary_lines.append(line_stripped)
                            if "Missing drawings:" in line_stripped:
                                self.capture_summary = False
                        
                        # Check for various error patterns
                        if any(error_indicator in line_stripped for error_indicator in [
                            "ERROR", "Error", "error",
                            "Traceback", "Exception",
                            "ValueError", "TypeError", "AttributeError", 
                            "UnicodeEncodeError", "UnicodeDecodeError",
                            "raise", "FAIL", "Failed", "failed"
                        ]):
                            script_has_errors = True
                            self.processing_errors.append(f"{script_name}: {line_stripped}")
                            self.log_message(line_stripped, "error")
                        elif "WARNING" in line_stripped or "Warning" in line_stripped:
                            self.log_message(line_stripped, "warning")
                        elif "SUCCESS" in line_stripped or "Successfully" in line_stripped:
                            self.log_message(line_stripped, "success")
                        
            process.wait()
            
            # Check return code and error status
            if process.returncode != 0 or script_has_errors:
                self.processing_errors.append(f"{script_name} exited with code {process.returncode}")
                return False
                
            return True
            
        except Exception as e:
            error_msg = f"Failed to run {script_name}: {str(e)}"
            self.log_message(error_msg, "error")
            self.terminal_output(f"\nERROR: {error_msg}\n")
            self.processing_errors.append(error_msg)
            return False
            
    def parse_script_output(self, line, start_progress, end_progress):
        """Parse script output for progress updates"""
        # Look for progress bar updates (format: |████████████████████| 100.0%)
        progress_match = re.search(r'\|[█\-]+\|\s*(\d+\.?\d*)%', line)
        if progress_match:
            script_progress = float(progress_match.group(1))
            # Map script progress to overall progress range
            overall_progress = start_progress + (script_progress / 100) * (end_progress - start_progress)
            self.update_progress(overall_progress)
            
    def update_progress(self, value, status=None):
        """Update progress bar and status label"""
        def update():
            self.progress_var.set(value)
            if status:
                self.progress_label.config(text=status)
                # Change color based on status
                if "Error" in status or "error" in status:
                    self.progress_label.config(fg=self.error_color)
                elif "complete" in status:
                    self.progress_label.config(fg=self.success_color)
                else:
                    self.progress_label.config(fg=self.primary_color)
                    
        self.root.after(0, update)
        
    def reset_controls(self):
        """Reset controls after processing"""
        self.is_processing = False
        self.sort_btn.config(state="normal", text="SORT PDFs")
        self.browse_btn.config(state="normal")
        self.csv_browse_btn.config(state="normal")


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = PDFSorterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()