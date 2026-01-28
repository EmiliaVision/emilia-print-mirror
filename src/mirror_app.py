"""
Spooler Queue Copy - Mirror Application

GUI application to configure and run the printer mirror service.
"""

import sys
import os
import time
import logging
import threading
import subprocess
from typing import Set, Optional, List

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

try:
    import win32print

    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QComboBox,
    QPushButton,
    QLabel,
    QTextEdit,
    QMessageBox,
    QCheckBox,
    QSpinBox,
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt6.QtGui import QFont, QTextCursor


class MirrorWorker(QThread):
    """Worker thread for the mirror service."""

    log_message = pyqtSignal(str)
    status_changed = pyqtSignal(str)
    job_copied = pyqtSignal(str, str, int)

    def __init__(self, source_printer: str, dest_printer: str, interval: float = 1.0):
        super().__init__()
        self.source_printer = source_printer
        self.dest_printer = dest_printer
        self.interval = interval
        self.running = False
        self.processed_jobs: Set[int] = set()
        self.spool_dir = os.path.join(
            os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS"
        )

    def log(self, message: str):
        self.log_message.emit(message)

    def _find_spool_file(self, job_id: int) -> Optional[str]:
        """Find the SPL file for a specific job."""
        try:
            patterns = [f"FP{job_id:05d}.SPL", f"{job_id:05d}.SPL"]

            for pattern in patterns:
                filepath = os.path.join(self.spool_dir, pattern)
                if os.path.exists(filepath):
                    return filepath

            spl_files = []
            for f in os.listdir(self.spool_dir):
                if f.upper().endswith(".SPL"):
                    full_path = os.path.join(self.spool_dir, f)
                    try:
                        mtime = os.path.getmtime(full_path)
                        size = os.path.getsize(full_path)
                        if size > 0:
                            spl_files.append((full_path, mtime))
                    except OSError:
                        continue

            if spl_files:
                spl_files.sort(key=lambda x: x[1], reverse=True)
                return spl_files[0][0]

        except Exception as e:
            self.log(f"Error finding spool: {e}")

        return None

    def _read_spool_data(self, job_id: int, retries: int = 5) -> Optional[bytes]:
        """Read spool data with retries."""
        for attempt in range(retries):
            spool_file = self._find_spool_file(job_id)

            if spool_file:
                try:
                    time.sleep(0.3)
                    with open(spool_file, "rb") as f:
                        data = f.read()
                    if data and len(data) > 0:
                        return data
                except Exception as e:
                    pass

            if attempt < retries - 1:
                time.sleep(0.5)

        return None

    def _copy_job(self, job_id: int, document_name: str) -> bool:
        """Copy a job to the destination printer."""
        try:
            spool_data = self._read_spool_data(job_id)

            if not spool_data:
                self.log(f"Could not read job {job_id}")
                return False

            handle = win32print.OpenPrinter(self.dest_printer)
            try:
                doc_info = (f"[MIRROR] {document_name}", "", "RAW")
                new_job_id = win32print.StartDocPrinter(handle, 1, doc_info)
                try:
                    win32print.StartPagePrinter(handle)
                    win32print.WritePrinter(handle, spool_data)
                    win32print.EndPagePrinter(handle)
                finally:
                    win32print.EndDocPrinter(handle)

                self.log(
                    f"OK! Job {job_id} -> {self.dest_printer} (ID: {new_job_id}, {len(spool_data)} bytes)"
                )
                self.job_copied.emit(self.source_printer, self.dest_printer, job_id)
                return True
            finally:
                win32print.ClosePrinter(handle)

        except Exception as e:
            self.log(f"Error copying job {job_id}: {e}")
            return False

    def _get_current_jobs(self) -> dict:
        """Get current jobs from source printer."""
        jobs = {}
        try:
            handle = win32print.OpenPrinter(self.source_printer)
            try:
                job_list = win32print.EnumJobs(handle, 0, -1, 1)
                for job in job_list:
                    job_id = job.get("JobId", 0)
                    doc_name = job.get("pDocument", "Unknown")
                    status = job.get("Status", 0)
                    jobs[job_id] = {
                        "document": doc_name,
                        "status": status,
                    }
            finally:
                win32print.ClosePrinter(handle)
        except Exception as e:
            pass

        return jobs

    def run(self):
        """Run the mirror service."""
        self.running = True
        self.status_changed.emit("running")

        self.log(f"Mirror started: {self.source_printer} -> {self.dest_printer}")

        # Verify spool access
        try:
            os.listdir(self.spool_dir)
        except PermissionError:
            self.log("ERROR: No access to spool. Run as Administrator.")
            self.status_changed.emit("error")
            return

        # Get existing jobs to ignore
        existing_jobs = self._get_current_jobs()
        self.processed_jobs = set(existing_jobs.keys())

        self.log("Waiting for print jobs...")

        while self.running:
            current_jobs = self._get_current_jobs()
            new_job_ids = set(current_jobs.keys()) - self.processed_jobs

            for job_id in sorted(new_job_ids):
                if not self.running:
                    break

                job_info = current_jobs[job_id]
                document = job_info["document"]

                if document.startswith("[MIRROR]") or document.startswith("[COPY]"):
                    self.processed_jobs.add(job_id)
                    continue

                self.log(f">>> New job: [{job_id}] {document}")
                time.sleep(1.0)

                self._copy_job(job_id, document)
                self.processed_jobs.add(job_id)

            self.processed_jobs &= set(current_jobs.keys())
            time.sleep(self.interval)

        self.status_changed.emit("stopped")
        self.log("Mirror stopped")

    def stop(self):
        self.running = False


class PrinterMirrorApp(QMainWindow):
    """Main Mirror Application."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.printers = []

        self._setup_ui()
        self._load_printers()

    def _setup_ui(self):
        self.setWindowTitle("Spooler Queue Copy - Printer Mirror")
        self.setMinimumSize(700, 500)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # === Configuration ===
        config_group = QGroupBox("Mirror Configuration")
        config_layout = QVBoxLayout(config_group)

        # Source printer
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("SOURCE Printer:"))
        self.source_combo = QComboBox()
        self.source_combo.setMinimumWidth(300)
        source_layout.addWidget(self.source_combo, 1)
        config_layout.addLayout(source_layout)

        # Arrow
        arrow_label = QLabel("↓  Jobs will be automatically copied to  ↓")
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setStyleSheet("font-weight: bold; color: #0066cc; padding: 10px;")
        config_layout.addWidget(arrow_label)

        # Destination printer
        dest_layout = QHBoxLayout()
        dest_layout.addWidget(QLabel("DESTINATION Printer:"))
        self.dest_combo = QComboBox()
        self.dest_combo.setMinimumWidth(300)
        dest_layout.addWidget(self.dest_combo, 1)
        config_layout.addLayout(dest_layout)

        # Options
        options_layout = QHBoxLayout()
        self.auto_config_check = QCheckBox("Auto-configure printer (KeepPrintedJobs)")
        self.auto_config_check.setChecked(True)
        options_layout.addWidget(self.auto_config_check)

        options_layout.addStretch()
        options_layout.addWidget(QLabel("Interval (sec):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 10)
        self.interval_spin.setValue(1)
        options_layout.addWidget(self.interval_spin)

        config_layout.addLayout(options_layout)
        layout.addWidget(config_group)

        # === Buttons ===
        btn_layout = QHBoxLayout()

        self.start_btn = QPushButton("▶ Start Mirror")
        self.start_btn.setMinimumHeight(40)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #218838; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.start_btn.clicked.connect(self._start_mirror)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⬛ Stop")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #c82333; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.stop_btn.clicked.connect(self._stop_mirror)
        btn_layout.addWidget(self.stop_btn)

        self.refresh_btn = QPushButton("🔄 Refresh Printers")
        self.refresh_btn.clicked.connect(self._load_printers)
        btn_layout.addWidget(self.refresh_btn)

        layout.addLayout(btn_layout)

        # === Status ===
        self.status_label = QLabel("Status: Stopped")
        self.status_label.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.status_label)

        # === Log ===
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout(log_group)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet("background-color: #1e1e1e; color: #00ff00;")
        log_layout.addWidget(self.log_text)

        clear_btn = QPushButton("Clear log")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)

        layout.addWidget(log_group, 1)

        # Initial message
        self._log("Spooler Queue Copy - Mirror v2.0")
        self._log("Select printers and press 'Start Mirror'")
        self._log("-" * 50)

    def _log(self, message: str):
        """Add message to log."""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.moveCursor(QTextCursor.MoveOperation.End)

    def _load_printers(self):
        """Load printer list."""
        self.source_combo.clear()
        self.dest_combo.clear()

        if not WINDOWS_AVAILABLE:
            self._log("ERROR: Requires Windows")
            return

        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = win32print.EnumPrinters(flags, None, 2)

            self.printers = []
            for p in printers:
                name = p["pPrinterName"]
                self.printers.append(name)
                self.source_combo.addItem(name)
                self.dest_combo.addItem(name)

            self._log(f"Found {len(self.printers)} printer(s)")

            # Select defaults if they exist
            for i, name in enumerate(self.printers):
                if "Org" in name:
                    self.source_combo.setCurrentIndex(i)
                if "Copy" in name:
                    self.dest_combo.setCurrentIndex(i)

        except Exception as e:
            self._log(f"Error loading printers: {e}")

    def _configure_printer(self, printer_name: str):
        """Configure printer to keep printed jobs."""
        try:
            cmd = f'Set-Printer -Name "{printer_name}" -KeepPrintedJobs $true'
            result = subprocess.run(
                ["powershell", "-Command", cmd], capture_output=True, text=True
            )
            if result.returncode == 0:
                self._log(f"Configured KeepPrintedJobs on {printer_name}")
                return True
            else:
                self._log(f"Warning: Could not configure KeepPrintedJobs")
                return False
        except Exception as e:
            self._log(f"Error configuring printer: {e}")
            return False

    def _start_mirror(self):
        """Start the mirror service."""
        source = self.source_combo.currentText()
        dest = self.dest_combo.currentText()

        if not source or not dest:
            QMessageBox.warning(self, "Error", "Select both printers")
            return

        if source == dest:
            QMessageBox.warning(self, "Error", "Printers must be different")
            return

        # Configure source printer if checked
        if self.auto_config_check.isChecked():
            self._configure_printer(source)

        # Start worker
        interval = self.interval_spin.value()
        self.worker = MirrorWorker(source, dest, interval)
        self.worker.log_message.connect(self._log)
        self.worker.status_changed.connect(self._on_status_changed)
        self.worker.start()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.source_combo.setEnabled(False)
        self.dest_combo.setEnabled(False)

        self._log(f"Starting mirror: {source} -> {dest}")

    def _stop_mirror(self):
        """Stop the mirror service."""
        if self.worker:
            self.worker.stop()
            self.worker.wait(3000)
            self.worker = None

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.source_combo.setEnabled(True)
        self.dest_combo.setEnabled(True)

    def _on_status_changed(self, status: str):
        """Handle status changes."""
        if status == "running":
            self.status_label.setText("Status: ● Running")
            self.status_label.setStyleSheet(
                "font-weight: bold; color: green; padding: 5px;"
            )
        elif status == "stopped":
            self.status_label.setText("Status: ⬛ Stopped")
            self.status_label.setStyleSheet(
                "font-weight: bold; color: gray; padding: 5px;"
            )
        elif status == "error":
            self.status_label.setText("Status: ✖ Error")
            self.status_label.setStyleSheet(
                "font-weight: bold; color: red; padding: 5px;"
            )

    def closeEvent(self, event):
        """Handle close event."""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirm",
                "Mirror is running. Stop and exit?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._stop_mirror()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Check admin
    if WINDOWS_AVAILABLE:
        import ctypes

        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0  # type: ignore
            if not is_admin:
                QMessageBox.warning(
                    None,
                    "Warning",
                    "It is recommended to run as Administrator\nto access spool files.",
                )
        except:
            pass

    window = PrinterMirrorApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
