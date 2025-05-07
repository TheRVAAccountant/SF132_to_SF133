import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time
from datetime import datetime
import threading
import queue
import logging
from typing import Callable, Optional
import os

class ExcelProcessorGUI:
    """
    A modern GUI interface for Excel file processing operations.
    Provides file selection, password input, and progress tracking functionality.
    """
    
    def __init__(self):
        """Initialize the GUI window and components."""
        self.root = tk.Tk()
        self.root.title("Excel File Processor")
        self.root.geometry("695x400")
        self.root.resizable(True, True)

        # Set window icon
        icon_path = "currency_icon.ico"
        self.root.iconbitmap(icon_path)

        # Set background color of the main window
        self.root.configure(background='#2e2e2e')

        # Apply forest-dark theme
        self.apply_forest_dark_theme()
        
        # Configure grid weight
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(4, weight=1)
        
        # Style configuration
        self.style = ttk.Style()
        self.style.configure('Custom.TFrame', background='#2e2e2e')
        self.style.configure('Custom.TButton', padding=5)
        self.style.configure('Custom.TLabel', background='#2e2e2e', foreground='#ffffff', padding=5)
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, style='Custom.TFrame', padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Initialize components
        self._init_file_selection()
        self._init_password_input()
        self._init_progress_components()
        self._init_status_components()
        
        # Processing variables
        self.start_time: Optional[float] = None
        self.processing_thread: Optional[threading.Thread] = None
        self.queue = queue.Queue()
    
    def apply_forest_dark_theme(self):
        """Apply the forest-dark theme to the GUI."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        tcl_file_path = os.path.join(script_dir, 'forest-dark.tcl')
        self.root.tk.call('source', tcl_file_path)
        ttk.Style().theme_use('forest-dark')

    def _init_file_selection(self):
        """Initialize file selection components."""
        # File selection frame
        file_frame = ttk.Frame(self.main_frame)
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5)
        
        # File path entry
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=0, column=0, padx=5, sticky="ew")

        # Browse button
        browse_btn = ttk.Button(
            file_frame,
            text="Browse",
            command=self._browse_file,
            style='Custom.TButton'
        )
        browse_btn.grid(row=0, column=1, padx=5)
        
        file_frame.grid_columnconfigure(0, weight=1)

    def _init_password_input(self):
        """Initialize password input components."""
        # Password frame
        pwd_frame = ttk.Frame(self.main_frame)
        pwd_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)
        
        # Password label and entry
        ttk.Label(
            pwd_frame,
            text="Sheet Password:",
            style='Custom.TLabel'
        ).grid(row=0, column=0, padx=5)
        
        self.password = tk.StringVar()
        self.pwd_entry = ttk.Entry(pwd_frame, textvariable=self.password, show="*")
        self.pwd_entry.grid(row=0, column=1, padx=5, sticky="ew")
        
        pwd_frame.grid_columnconfigure(1, weight=1)
        
    def _init_progress_components(self):
        """Initialize progress bar and related components."""
        # Progress frame
        progress_frame = ttk.Frame(self.main_frame)
        progress_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=5)
        
        # Process button
        self.process_btn = ttk.Button(
            progress_frame,
            text="Process File",
            command=self._start_processing,
            style='Custom.TButton'
        )
        self.process_btn.grid(row=1, column=0, pady=10)
        
        progress_frame.grid_columnconfigure(0, weight=1)
        
    def _init_status_components(self):
        """Initialize status display components."""
        # Status frame
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        
        # Status text
        self.status_text = tk.Text(status_frame, height=10, wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky="ew")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        status_frame.grid_columnconfigure(0, weight=1)
        
    def _browse_file(self):
        """Open file dialog for selecting Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            
    def update_progress(self, value: float, message: str):
        """Update progress bar and status message."""
        self.progress_var.set(value)
        self.update_status(message)
        
    def update_status(self, message: str):
        """Update status text with timestamp."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        
    def _start_processing(self):
        """Start the file processing in a separate thread."""
        if not self.file_path.get():
            messagebox.showwarning("Input Required", "Please select a file first.")
            return
            
        if not self.password.get():
            messagebox.showwarning("Input Required", "Please enter the sheet password.")
            return
        
        # Check if file exists
        if not os.path.exists(self.file_path.get()):
            messagebox.showerror("File Error", f"File not found: {self.file_path.get()}")
            return
            
        self.start_time = time.time()
        self.process_btn.configure(state="disabled")
        self.progress_var.set(0)
        
        # Clear status text
        self.status_text.delete(1.0, tk.END)
        self.update_status("Processing started...")
        
        # Start processing thread
        self.processing_thread = threading.Thread(
            target=self._processing_worker,
            daemon=True
        )
        self.processing_thread.start()
        
        # Start progress monitoring
        self.root.after(100, self._check_progress)
        
    def _processing_worker(self):
        """Worker function for processing the Excel file."""
        try:
            # This will be replaced with the actual processing function
            if self.process_callback:
                self.process_callback(
                    self.file_path.get(),
                    self.password.get(),
                    self.queue
                )
        except Exception as e:
            self.queue.put(("error", str(e)))
        finally:
            self.queue.put(("done", None))
            
    def _check_progress(self):
        """Check progress queue and update GUI."""
        try:
            while True:
                msg_type, data = self.queue.get_nowait()
                
                if msg_type == "progress":
                    value, message = data
                    self.update_progress(value, message)
                elif msg_type == "status":
                    self.update_status(data)
                elif msg_type == "error":
                    self.update_status(f"Error: {data}")
                    messagebox.showerror("Processing Error", data)
                    self._finish_processing()
                elif msg_type == "warning":
                    self.update_status(f"Warning: {data}")
                    messagebox.showwarning("Processing Warning", data)
                elif msg_type == "success":
                    self.update_status(f"Success: {data}")
                    messagebox.showinfo("Processing Complete", data)
                    self._finish_processing()
                    return
                elif msg_type == "done":
                    self._finish_processing()
                    return
                    
                self.queue.task_done()
                
        except queue.Empty:
            if self.processing_thread and self.processing_thread.is_alive():
                self.root.after(100, self._check_progress)
                
    def _finish_processing(self):
        """Clean up after processing is complete."""
        elapsed_time = time.time() - self.start_time
        self.update_status(f"Processing completed in {elapsed_time:.2f} seconds")
        self.process_btn.configure(state="normal")
        
    def set_process_callback(self, callback: Callable):
        """Set the callback function for file processing."""
        self.process_callback = callback
        
    def run(self):
        """Start the GUI application."""
        self.root.mainloop()