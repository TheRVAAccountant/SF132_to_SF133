"""
GUI interface for Excel processing application.

This module provides a modern Tkinter-based GUI for the Excel processing
application with file selection, progress tracking, and status display.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time
from datetime import datetime
import threading
import queue
import logging
import sys
import os
from typing import Callable, Optional, Dict, Any

# Type aliases
ProcessCallback = Callable[[str, str, queue.Queue], bool]

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

        # Set window icon if available
        icon_path = "currency_icon.ico"
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError:
                # On some platforms iconbitmap might not work
                pass

        # Set up Windows integration if running on Windows
        if sys.platform.startswith('win'):
            try:
                from ..utils.win_path_handler import normalize_windows_path
                self._normalize_path = normalize_windows_path
                logging.info("Windows path normalization enabled for GUI")
            except ImportError:
                self._normalize_path = lambda x: x
                logging.warning("Windows path normalization not available")

        # Set background color of the main window
        self.root.configure(background='#2e2e2e')

        # Apply theme
        self._apply_theme()
        
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
        
        # Initialize menu
        self._init_menu()
        
        # Initialize components
        self._init_file_selection()
        self._init_password_input()
        self._init_progress_components()
        self._init_status_components()
        
        # Processing variables
        self.start_time: Optional[float] = None
        self.processing_thread: Optional[threading.Thread] = None
        self.queue = queue.Queue()
        self.process_callback: Optional[ProcessCallback] = None
        
        # Recovery mode flag - enabled by default on Windows
        self.recovery_mode = tk.BooleanVar(value=sys.platform.startswith('win'))
    
    def _apply_theme(self):
        """Apply the theme to the GUI."""
        # Try to use the forest theme if available
        try:
            script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            tcl_file_path = os.path.join(script_dir, 'forest-dark.tcl')
            if os.path.exists(tcl_file_path):
                self.root.tk.call('source', tcl_file_path)
                ttk.Style().theme_use('forest-dark')
            else:
                # Fall back to a built-in theme
                ttk.Style().theme_use('clam')
        except Exception as e:
            logging.warning(f"Error applying theme: {e}")
            # Fall back to a built-in theme
            ttk.Style().theme_use('clam')

    def _init_menu(self):
        """Initialize the application menu."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel File", command=self._browse_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Options menu
        options_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Options", menu=options_menu)
        
        # Add recovery mode option (Windows only)
        if sys.platform.startswith('win'):
            options_menu.add_checkbutton(
                label="Enable Recovery Mode",
                variable=self.recovery_mode,
                onvalue=True,
                offvalue=False
            )
            
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)
        
        # Add Windows-specific menu items if on Windows
        if sys.platform.startswith('win'):
            tools_menu = tk.Menu(menubar, tearoff=0)
            menubar.add_cascade(label="Tools", menu=tools_menu)
            tools_menu.add_command(label="Fix Excel File", command=self._fix_excel_file)
            tools_menu.add_command(label="Close Excel Instances", command=self._close_excel_instances)

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

    def _show_about(self):
        """Show about dialog."""
        messagebox.showinfo(
            "About Excel Processor",
            "SF132 to SF133 Excel File Processor\n\n"
            "A tool for processing SF132 and SF133 reconciliation Excel files.\n\n"
            "Version: 1.0"
        )
        
    def _fix_excel_file(self):
        """Fix an Excel file with access issues."""
        if not sys.platform.startswith('win'):
            messagebox.showinfo("Not Available", "This feature is only available on Windows.")
            return
            
        filename = filedialog.askopenfilename(
            title="Select Excel File to Fix",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not filename:
            return
        
        # Normalize path if on Windows
        filename = self._normalize_path(filename)
            
        try:
            from ..modules.excel_recovery import fix_file_in_use_error
            
            self.update_status(f"Attempting to fix Excel file: {filename}...")
            success, result_path = fix_file_in_use_error(filename)
            
            if success:
                self.update_status(f"Successfully fixed file: {result_path}")
                messagebox.showinfo("Fix Complete", f"File has been fixed and saved to:\n{result_path}")
                
                # Offer to load the fixed file
                if messagebox.askyesno("Load Fixed File", "Would you like to load the fixed file?"):
                    self.file_path.set(result_path)
            else:
                self.update_status(f"Failed to fix file: {result_path}")
                messagebox.showerror("Fix Failed", f"Failed to fix file: {result_path}")
                
        except ImportError:
            messagebox.showerror("Not Available", "Excel recovery modules are not available.")
            
    def _close_excel_instances(self):
        """Close all running Excel instances."""
        if not sys.platform.startswith('win'):
            messagebox.showinfo("Not Available", "This feature is only available on Windows.")
            return
            
        try:
            from ..modules.excel_handler import close_excel_instances
            
            self.update_status("Closing Excel instances...")
            terminated = close_excel_instances()
            
            if terminated:
                self.update_status(f"Closed {len(terminated)} Excel instance(s)")
                messagebox.showinfo("Excel Closed", f"Successfully closed {len(terminated)} Excel instance(s).")
            else:
                self.update_status("No Excel instances found to close")
                messagebox.showinfo("Excel Instances", "No Excel instances were found running.")
                
        except ImportError:
            messagebox.showerror("Not Available", "Excel handler module is not available.")
        
    def _browse_file(self):
        """Open file dialog for selecting Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            # Normalize Windows path if running on Windows
            filename = self._normalize_path(filename)
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
            # Check for Windows-specific file handling
            if sys.platform.startswith('win'):
                try:
                    # Import Windows integration tools if available
                    from ..modules.excel_handler import close_excel_instances
                    # Close any Excel instances before processing
                    close_excel_instances()
                    self.queue.put(("status", "Closed any running Excel instances for clean processing"))
                except ImportError:
                    self.queue.put(("status", "Windows integration modules not available"))
            
            # Call the actual processing function
            if self.process_callback:
                # Check if recovery mode should be attempted for Windows
                if sys.platform.startswith('win') and self.recovery_mode.get():
                    try:
                        from ..modules.excel_recovery import process_with_recovery
                        # Use recovery-enhanced processing
                        self.queue.put(("status", "Using Windows file recovery enhancements"))
                        success, result = process_with_recovery(
                            self.process_callback,
                            self.file_path.get(),
                            self.password.get(),
                            self.queue
                        )
                        if not success:
                            self.queue.put(("error", result))
                    except ImportError:
                        # Fall back to regular processing
                        self.process_callback(
                            self.file_path.get(),
                            self.password.get(),
                            self.queue
                        )
                else:
                    # Regular processing for non-Windows platforms or recovery mode disabled
                    self.process_callback(
                        self.file_path.get(),
                        self.password.get(),
                        self.queue
                    )
        except Exception as e:
            self.queue.put(("error", str(e)))
        finally:
            # Clean up resources if running on Windows
            if sys.platform.startswith('win'):
                try:
                    from ..modules.file_operations import cleanup_temp_files
                    cleanup_temp_files()
                    self.queue.put(("status", "Cleaned up temporary files"))
                except ImportError:
                    pass
                    
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
        
    def set_process_callback(self, callback: ProcessCallback):
        """Set the callback function for file processing."""
        self.process_callback = callback
        
    def run(self):
        """Start the GUI application."""
        self.root.mainloop()