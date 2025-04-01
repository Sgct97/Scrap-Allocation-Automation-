import sys
import os
import subprocess
import threading
from pathlib import Path
import traceback

# Try importing PySide6 components
try:
    from PySide6.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QLabel, QTextEdit, QFileDialog, QSizePolicy
    )
    from PySide6.QtCore import Qt, Signal, QObject, Slot
    PYSIDE6_AVAILABLE = True
except ImportError:
    print("Error: PySide6 library not found. Please install it using:")
    print("pip3 install PySide6")
    PYSIDE6_AVAILABLE = False
    # Define dummy classes if PySide6 is not available to avoid syntax errors later
    class QObject: pass
    class QWidget: pass

# Determine if running as a bundled executable (PyInstaller) or a script
if getattr(sys, 'frozen', False):
    BASE_PATH = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(os.path.dirname(sys.executable))
    ALLOCATOR_SCRIPT_PATH = BASE_PATH / "scrap_allocator.py"
else:
    BASE_PATH = Path(__file__).parent
    ALLOCATOR_SCRIPT_PATH = BASE_PATH / "scrap_allocator.py"

# --- Constants ---
WINDOW_TITLE = "Scrap Allocation Runner (Qt)"
WINDOW_WIDTH = 550
WINDOW_HEIGHT = 350
SUCCESS_COLOR = "green"
ERROR_COLOR = "red"
STATUS_COLOR = "gray"
WORKING_COLOR = "orange"

# --- Worker Signal Class ---
# Needed to safely update the GUI from the worker thread
class WorkerSignals(QObject):
    finished = Signal(int, str, str) # return_code, stdout, stderr
    status_update = Signal(str, str) # message, color
    error = Signal(str)              # error message

# --- Worker Thread Class ---
class AllocationWorker(QObject):
    def __init__(self, command, parent=None):
        super().__init__(parent)
        self.command = command
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        try:
            self.signals.status_update.emit("Starting allocation script...", WORKING_COLOR)
            print(f"Running command: {' '.join(self.command)}")

            startupinfo = None
            if os.name == 'nt': # Windows
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            process = subprocess.run(
                self.command,
                capture_output=True,
                text=True,
                check=False,
                startupinfo=startupinfo,
                encoding='utf-8'
            )

            print(f"Subprocess finished. Return code: {process.returncode}")
            # Limit output length slightly for display if needed
            stdout_short = (process.stdout or "")[:2000]
            stderr_short = (process.stderr or "")[:2000]
            if len(process.stdout or "") > 2000: stdout_short += "\n... (output truncated)"
            if len(process.stderr or "") > 2000: stderr_short += "\n... (output truncated)"

            self.signals.finished.emit(process.returncode, stdout_short, stderr_short)

        except FileNotFoundError:
             self.signals.error.emit(f"Error: Python executable or allocator script not found.\nCommand: {' '.join(self.command)}")
        except Exception as e:
            error_traceback = traceback.format_exc()
            print(f"Worker Error during subprocess run:\n{error_traceback}")
            self.signals.error.emit(f"An unexpected error occurred in the worker thread:\n{e}")
        finally:
            # Explicitly tell the QThread managing this worker to quit its event loop
            if self.thread() is not None: # Check if thread exists
                 print("Worker telling thread to quit.") # Debug print
                 self.thread().quit()


# --- Main Application Window ---
class AppWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worksheet_path = None
        self.recap_path = None
        self.worker_thread = None # To hold the thread reference
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(WINDOW_TITLE)
        self.setMinimumSize(WINDOW_WIDTH, WINDOW_HEIGHT)

        # --- Layouts ---
        main_layout = QVBoxLayout(self)
        file_layout = QVBoxLayout() # Vertical layout for file sections
        ws_layout = QHBoxLayout()   # Horizontal for worksheet button+label
        recap_layout = QHBoxLayout() # Horizontal for recap button+label
        status_layout = QVBoxLayout()
        button_layout = QHBoxLayout()

        # --- Widgets ---
        # Worksheet Selection
        self.ws_button = QPushButton("Select Worksheet...")
        self.ws_button.clicked.connect(self.browse_worksheet)
        self.ws_path_label = QLabel("No file selected")
        self.ws_path_label.setWordWrap(True)
        ws_layout.addWidget(self.ws_button)
        ws_layout.addWidget(self.ws_path_label, 1) # Label takes expanding space

        # Recap Selection
        self.recap_button = QPushButton("Select Recap File...")
        self.recap_button.clicked.connect(self.browse_recap)
        self.recap_path_label = QLabel("No file selected")
        self.recap_path_label.setWordWrap(True)
        recap_layout.addWidget(self.recap_button)
        recap_layout.addWidget(self.recap_path_label, 1)

        # Status Area
        self.status_textbox = QTextEdit()
        self.status_textbox.setReadOnly(True)
        self.status_textbox.setPlaceholderText("Status: Waiting for files...")
        status_layout.addWidget(QLabel("Status:"))
        status_layout.addWidget(self.status_textbox)

        # Run Button
        self.run_button = QPushButton("Run Allocation")
        self.run_button.clicked.connect(self.run_allocation_thread)
        self.run_button.setEnabled(False) # Disabled initially
        button_layout.addStretch() # Center button (optional)
        button_layout.addWidget(self.run_button)
        button_layout.addStretch()

        # --- Assemble Layouts ---
        file_layout.addLayout(ws_layout)
        file_layout.addLayout(recap_layout)

        main_layout.addLayout(file_layout)
        main_layout.addLayout(status_layout)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

    def browse_file(self, file_type):
        initial_dir = str(Path.home())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select {file_type.capitalize()} File",
            initial_dir,
            "Excel Files (*.xlsx);;All Files (*)"
        )

        if not file_path: # User cancelled
            return

        dropped_path = Path(file_path)
        filename = str(dropped_path) # Display full path

        if file_type == 'worksheet':
            self.worksheet_path = str(dropped_path)
            self.ws_path_label.setText(filename) # Update label
            self.update_status("Worksheet file selected.")
            print(f"Worksheet path set: {self.worksheet_path}")
        elif file_type == 'recap':
            self.recap_path = str(dropped_path)
            self.recap_path_label.setText(filename) # Update label
            self.update_status("Recap file selected.")
            print(f"Recap path set: {self.recap_path}")

        if self.worksheet_path and self.recap_path:
            self.run_button.setEnabled(True)
            self.update_status("Both files selected. Ready to run.")

    def browse_worksheet(self):
        self.browse_file('worksheet')

    def browse_recap(self):
        self.browse_file('recap')

    # Use signals/slots for safe UI updates from thread
    @Slot(str, str)
    def update_status(self, message, color=STATUS_COLOR):
        self.status_textbox.clear() # Clear previous messages
        self.status_textbox.setTextColor(color) # Set color
        self.status_textbox.append(message) # Append new message
        self.status_textbox.setTextColor("black") # Reset to default if needed

    @Slot(int, str, str)
    def handle_worker_finished(self, return_code, stdout, stderr):
        # Always print full output to terminal for debugging
        print(f"Subprocess finished. Return code: {return_code}")
        print(f"Subprocess stdout:\n{stdout}")
        print(f"Subprocess stderr:\n{stderr}")
        
        if return_code == 0:
            # output_message = "Allocation finished successfully!\n\n--- Script Output ---\n" + stdout # Old way
            self.update_status("Allocation finished successfully!", SUCCESS_COLOR)
        else:
            # error_message = f"Error running allocation script (Exit Code: {return_code}).\n\n--- Script Output ---\n{stdout}\n\n--- Error Output ---\n{stderr}" # Old way
            # Provide a simpler error in GUI, full details are in terminal
            error_summary = stderr.splitlines()[-1] if stderr else f"Exit Code: {return_code}" # Try to get last line of error
            self.update_status(f"Allocation failed: {error_summary}\n(See terminal for full details)", ERROR_COLOR)

        # Re-enable buttons
        self.run_button.setEnabled(True)
        self.ws_button.setEnabled(True)
        self.recap_button.setEnabled(True)
        self.worker_thread = None # Clear thread reference

    @Slot(str)
    def handle_worker_error(self, error_message):
        # Keep GUI error message relatively concise
        print(f"GUI Worker Error Slot Received: {error_message}") # Log full error to terminal
        self.update_status(f"GUI Error: {error_message}\n(See terminal for full traceback)", ERROR_COLOR)
        # Re-enable buttons
        self.run_button.setEnabled(True)
        self.ws_button.setEnabled(True)
        self.recap_button.setEnabled(True)
        self.worker_thread = None

    def run_allocation_thread(self):
        if not self.worksheet_path or not self.recap_path:
            self.update_status("Error: Both worksheet and recap files must be selected.", ERROR_COLOR)
            return
        if not ALLOCATOR_SCRIPT_PATH.is_file():
             self.update_status(f"Error: Allocator script not found at expected location:\n{ALLOCATOR_SCRIPT_PATH}", ERROR_COLOR)
             return

        self.run_button.setEnabled(False)
        self.ws_button.setEnabled(False)
        self.recap_button.setEnabled(False)
        # self.update_status("Starting allocation script...", WORKING_COLOR) # Status updated by worker signal

        command = [
            sys.executable,
            str(ALLOCATOR_SCRIPT_PATH),
            self.worksheet_path,
            self.recap_path
        ]

        # --- Setup Threading ---
        # Need to use QThread for proper integration with Qt event loop
        from PySide6.QtCore import QThread # Import QThread here

        self.worker_thread = QThread()
        self.worker = AllocationWorker(command)
        self.worker.moveToThread(self.worker_thread)

        # Connect signals
        self.worker.signals.finished.connect(self.handle_worker_finished)
        self.worker.signals.status_update.connect(self.update_status)
        self.worker.signals.error.connect(self.handle_worker_error)
        self.worker_thread.started.connect(self.worker.run)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater) # Clean up thread

        self.worker_thread.start()

# --- Run the App ---
if __name__ == "__main__":
    if not PYSIDE6_AVAILABLE:
        sys.exit(1) # Exit if import failed

    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec()) 