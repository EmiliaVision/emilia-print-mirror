"""
Emilia Print Mirror - Print Queue Mirroring Application

GUI application to configure and run the printer mirror service.
Supports multiple source printers mirroring to one destination.
"""

import sys
import os
import time
import json
import logging
import subprocess
import base64
from typing import Set, Optional, List, Dict
from pathlib import Path

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

try:
    import pywintypes  # noqa: F401 - Required for PyInstaller
    import pythoncom  # noqa: F401 - Required for PyInstaller
    import win32api  # noqa: F401 - Required for PyInstaller
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
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QByteArray
from PyQt6.QtGui import QFont, QTextCursor, QIcon, QPixmap
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtCore import QBuffer


# Emilia Flower Icon (PiFlower from Phosphor Icons) - Pink color
FLOWER_ICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256" fill="#E91E63">
  <path d="M210.35,129.36c-.81-.47-1.7-.92-2.62-1.36.92-.44,1.81-.89,2.62-1.36a40,40,0,1,0-40-69.28c-.81.47-1.65,1-2.48,1.59.08-1,.13-2,.13-3a40,40,0,0,0-80,0c0,.94,0,1.94.13,3-.83-.57-1.67-1.12-2.48-1.59a40,40,0,1,0-40,69.28c.81.47,1.7.92,2.62,1.36-.92.44-1.81.89-2.62,1.36a40,40,0,1,0,40,69.28c.81-.47,1.65-1,2.48-1.59-.08,1-.13,2-.13,2.95a40,40,0,0,0,80,0c0-.94-.05-1.94-.13-2.95.83.57,1.67,1.12,2.48,1.59A39.79,39.79,0,0,0,190.29,204a40.43,40.43,0,0,0,10.42-1.38,40,40,0,0,0,9.64-73.28ZM104,128a24,24,0,1,1,24,24A24,24,0,0,1,104,128Zm74.35-56.79a24,24,0,1,1,24,41.57c-6.27,3.63-18.61,6.13-35.16,7.19A40,40,0,0,0,154.53,98.1C163.73,84.28,172.08,74.84,178.35,71.21ZM128,32a24,24,0,0,1,24,24c0,7.24-4,19.19-11.36,34.06a39.81,39.81,0,0,0-25.28,0C108,75.19,104,63.24,104,56A24,24,0,0,1,128,32ZM44.86,80a24,24,0,0,1,32.79-8.79c6.27,3.63,14.62,13.07,23.82,26.89A40,40,0,0,0,88.81,120c-16.55-1.06-28.89-3.56-35.16-7.18A24,24,0,0,1,44.86,80ZM77.65,184.79a24,24,0,1,1-24-41.57c6.27-3.63,18.61-6.13,35.16-7.19a40,40,0,0,0,12.66,21.87C92.27,171.72,83.92,181.16,77.65,184.79ZM128,224a24,24,0,0,1-24-24c0-7.24,4-19.19,11.36-34.06a39.81,39.81,0,0,0,25.28,0C148,180.81,152,192.76,152,200A24,24,0,0,1,128,224Zm83.14-48a24,24,0,0,1-32.79,8.79c-6.27-3.63-14.62-13.07-23.82-26.89A40,40,0,0,0,167.19,136c16.55,1.06,28.89,3.56,35.16,7.18A24,24,0,0,1,211.14,176Z"/>
</svg>"""

APP_NAME = "Emilia Print Mirror"
APP_VERSION = "2.2.0"

# Configuration file path
CONFIG_PATH = (
    Path(os.environ.get("APPDATA", os.path.expanduser("~")))
    / "EmiliaPrintMirror"
    / "gui_config.json"
)


def load_config() -> dict:
    """Load GUI configuration from file."""
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(config: dict):
    """Save GUI configuration to file."""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


def get_app_icon() -> QIcon:
    """Create QIcon from embedded SVG."""
    svg_bytes = QByteArray(FLOWER_ICON_SVG.encode())
    renderer = QSvgRenderer(svg_bytes)

    # Create pixmap at different sizes for the icon
    icon = QIcon()
    for size in [16, 32, 48, 64, 128, 256]:
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        from PyQt6.QtGui import QPainter

        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        icon.addPixmap(pixmap)

    return icon


class MirrorWorker(QThread):
    """Worker thread for the mirror service with multiple source support."""

    log_message = pyqtSignal(str)
    status_changed = pyqtSignal(str)
    job_copied = pyqtSignal(str, str, int)

    def __init__(
        self, source_printers: List[str], dest_printer: str, interval: float = 1.0
    ):
        super().__init__()
        self.source_printers = source_printers
        self.dest_printer = dest_printer
        self.interval = interval
        self.running = False
        self.processed_jobs: Dict[str, Set[int]] = {p: set() for p in source_printers}
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
                except Exception:
                    pass

            if attempt < retries - 1:
                time.sleep(0.5)

        return None

    def _copy_job(self, source_printer: str, job_id: int, document_name: str) -> bool:
        """Copy a job to the destination printer."""
        try:
            spool_data = self._read_spool_data(job_id)

            if not spool_data:
                self.log(f"Could not read job {job_id}")
                return False

            handle = win32print.OpenPrinter(self.dest_printer)
            try:
                doc_info = (f"[MIRROR:{source_printer}] {document_name}", "", "RAW")
                new_job_id = win32print.StartDocPrinter(handle, 1, doc_info)
                try:
                    win32print.StartPagePrinter(handle)
                    win32print.WritePrinter(handle, spool_data)
                    win32print.EndPagePrinter(handle)
                finally:
                    win32print.EndDocPrinter(handle)

                self.log(
                    f"OK! [{source_printer}] Job {job_id} -> {self.dest_printer} (ID: {new_job_id}, {len(spool_data)} bytes)"
                )
                self.job_copied.emit(source_printer, self.dest_printer, job_id)
                return True
            finally:
                win32print.ClosePrinter(handle)

        except Exception as e:
            self.log(f"Error copying job {job_id}: {e}")
            return False

    def _get_current_jobs(self, printer_name: str) -> dict:
        """Get current jobs from a printer."""
        jobs = {}
        try:
            handle = win32print.OpenPrinter(printer_name)
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
        except Exception:
            pass

        return jobs

    def run(self):
        """Run the mirror service for multiple source printers."""
        self.running = True
        self.status_changed.emit("running")

        sources_str = ", ".join(self.source_printers)
        self.log(f"Mirror started: [{sources_str}] -> {self.dest_printer}")

        try:
            os.listdir(self.spool_dir)
        except PermissionError:
            self.log("ERROR: No access to spool. Run as Administrator.")
            self.status_changed.emit("error")
            return

        for printer in self.source_printers:
            existing_jobs = self._get_current_jobs(printer)
            self.processed_jobs[printer] = set(existing_jobs.keys())
            self.log(f"[{printer}] Ignoring {len(existing_jobs)} existing job(s)")

        self.log("Waiting for print jobs...")

        while self.running:
            for printer in self.source_printers:
                if not self.running:
                    break

                current_jobs = self._get_current_jobs(printer)
                new_job_ids = set(current_jobs.keys()) - self.processed_jobs[printer]

                for job_id in sorted(new_job_ids):
                    if not self.running:
                        break

                    job_info = current_jobs[job_id]
                    document = job_info["document"]

                    if document.startswith("[MIRROR"):
                        self.processed_jobs[printer].add(job_id)
                        continue

                    self.log(f">>> [{printer}] New job: [{job_id}] {document}")
                    time.sleep(1.0)

                    self._copy_job(printer, job_id, document)
                    self.processed_jobs[printer].add(job_id)

                self.processed_jobs[printer] &= set(current_jobs.keys())

            time.sleep(self.interval)

        self.status_changed.emit("stopped")
        self.log("Mirror stopped")

    def stop(self):
        self.running = False


class PrinterMirrorApp(QMainWindow):
    """Main Emilia Print Mirror Application."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.printers = []
        self.config = load_config()

        self._setup_ui()
        self._load_printers()
        self._apply_saved_config()

        # Auto-start if configured
        if self.config.get("auto_start", False):
            # Use a timer to start after UI is ready
            from PyQt6.QtCore import QTimer

            QTimer.singleShot(500, self._auto_start_mirror)

    def _setup_ui(self):
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(800, 650)

        # Set application icon
        app_icon = get_app_icon()
        self.setWindowIcon(app_icon)
        QApplication.instance().setWindowIcon(app_icon)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # === Header with logo ===
        header_layout = QHBoxLayout()

        # Logo
        logo_label = QLabel()
        svg_bytes = QByteArray(FLOWER_ICON_SVG.encode())
        renderer = QSvgRenderer(svg_bytes)
        pixmap = QPixmap(48, 48)
        pixmap.fill(Qt.GlobalColor.transparent)
        from PyQt6.QtGui import QPainter

        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        logo_label.setPixmap(pixmap)
        header_layout.addWidget(logo_label)

        # Title
        title_label = QLabel(f"<h2 style='color: #E91E63; margin: 0;'>{APP_NAME}</h2>")
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        layout.addLayout(header_layout)

        # === Configuration ===
        config_group = QGroupBox("Mirror Configuration")
        config_layout = QVBoxLayout(config_group)

        # Source printers (multi-select list)
        source_layout = QVBoxLayout()
        source_label = QLabel("SOURCE Printers (select one or more):")
        source_label.setStyleSheet("font-weight: bold;")
        source_layout.addWidget(source_label)

        self.source_list = QListWidget()
        self.source_list.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection
        )
        self.source_list.setMaximumHeight(120)
        source_layout.addWidget(self.source_list)

        config_layout.addLayout(source_layout)

        # Arrow
        arrow_label = QLabel("↓  Jobs from selected printers will be copied to  ↓")
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setStyleSheet("font-weight: bold; color: #E91E63; padding: 10px;")
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
        self.auto_config_check = QCheckBox("Auto-configure printers (KeepPrintedJobs)")
        self.auto_config_check.setChecked(True)
        options_layout.addWidget(self.auto_config_check)

        options_layout.addStretch()
        options_layout.addWidget(QLabel("Interval (sec):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 10)
        self.interval_spin.setValue(1)
        options_layout.addWidget(self.interval_spin)

        config_layout.addLayout(options_layout)

        # Auto-start option
        autostart_layout = QHBoxLayout()
        self.auto_start_check = QCheckBox("Auto-start mirror when application opens")
        self.auto_start_check.setToolTip(
            "Automatically start mirroring when the application launches"
        )
        autostart_layout.addWidget(self.auto_start_check)
        autostart_layout.addStretch()
        config_layout.addLayout(autostart_layout)

        layout.addWidget(config_group)

        # === Buttons ===
        btn_layout = QHBoxLayout()

        self.start_btn = QPushButton("▶ Start Mirror")
        self.start_btn.setMinimumHeight(40)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #E91E63;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #C2185B; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.start_btn.clicked.connect(self._start_mirror)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⬛ Stop")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #424242;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #212121; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.stop_btn.clicked.connect(self._stop_mirror)
        btn_layout.addWidget(self.stop_btn)

        self.refresh_btn = QPushButton("🔄 Refresh")
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
        self.log_text.setStyleSheet("background-color: #1a1a2e; color: #E91E63;")
        log_layout.addWidget(self.log_text)

        clear_btn = QPushButton("Clear log")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)

        layout.addWidget(log_group, 1)

        # Initial message
        self._log(f"🌸 {APP_NAME} v{APP_VERSION}")
        self._log("Select source printer(s) and destination, then press 'Start Mirror'")
        self._log("-" * 60)

    def _log(self, message: str):
        """Add message to log."""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.moveCursor(QTextCursor.MoveOperation.End)

    def _load_printers(self):
        """Load printer list."""
        self.source_list.clear()
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

                item = QListWidgetItem(name)
                self.source_list.addItem(item)
                self.dest_combo.addItem(name)

            self._log(f"Found {len(self.printers)} printer(s)")

            # Auto-select printers with "Org" in name as sources
            for i in range(self.source_list.count()):
                item = self.source_list.item(i)
                if item and "Org" in item.text():
                    item.setSelected(True)

            # Auto-select printer with "Copy" as destination
            for i, name in enumerate(self.printers):
                if "Copy" in name:
                    self.dest_combo.setCurrentIndex(i)
                    break

        except Exception as e:
            self._log(f"Error loading printers: {e}")

    def _apply_saved_config(self):
        """Apply saved configuration to UI."""
        if not self.config:
            return

        # Apply saved source printers
        saved_sources = self.config.get("source_printers", [])
        if saved_sources:
            for i in range(self.source_list.count()):
                item = self.source_list.item(i)
                if item:
                    item.setSelected(item.text() in saved_sources)

        # Apply saved destination
        saved_dest = self.config.get("dest_printer", "")
        if saved_dest:
            index = self.dest_combo.findText(saved_dest)
            if index >= 0:
                self.dest_combo.setCurrentIndex(index)

        # Apply saved interval
        saved_interval = self.config.get("interval", 1)
        self.interval_spin.setValue(saved_interval)

        # Apply auto-start setting
        self.auto_start_check.setChecked(self.config.get("auto_start", False))

        self._log("Loaded saved configuration")

    def _save_current_config(self):
        """Save current configuration to file."""
        self.config = {
            "source_printers": self._get_selected_sources(),
            "dest_printer": self.dest_combo.currentText(),
            "interval": self.interval_spin.value(),
            "auto_start": self.auto_start_check.isChecked(),
        }
        save_config(self.config)
        self._log("Configuration saved")

    def _auto_start_mirror(self):
        """Auto-start mirror if configured."""
        sources = self._get_selected_sources()
        dest = self.dest_combo.currentText()

        if sources and dest and dest not in sources:
            self._log("Auto-starting mirror...")
            self._start_mirror()
        else:
            self._log("Auto-start skipped: invalid configuration")

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
                self._log(f"Warning: Could not configure {printer_name}")
                return False
        except Exception as e:
            self._log(f"Error configuring printer: {e}")
            return False

    def _get_selected_sources(self) -> List[str]:
        """Get list of selected source printers."""
        selected = []
        for item in self.source_list.selectedItems():
            selected.append(item.text())
        return selected

    def _start_mirror(self):
        """Start the mirror service."""
        sources = self._get_selected_sources()
        dest = self.dest_combo.currentText()

        if not sources:
            QMessageBox.warning(self, "Error", "Select at least one source printer")
            return

        if not dest:
            QMessageBox.warning(self, "Error", "Select a destination printer")
            return

        if dest in sources:
            QMessageBox.warning(self, "Error", "Destination cannot be a source printer")
            return

        # Save configuration when starting
        self._save_current_config()

        if self.auto_config_check.isChecked():
            for source in sources:
                self._configure_printer(source)

        interval = self.interval_spin.value()
        self.worker = MirrorWorker(sources, dest, interval)
        self.worker.log_message.connect(self._log)
        self.worker.status_changed.connect(self._on_status_changed)
        self.worker.start()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.source_list.setEnabled(False)
        self.dest_combo.setEnabled(False)

        self._log(f"Starting mirror: {sources} -> {dest}")

    def _stop_mirror(self):
        """Stop the mirror service."""
        if self.worker:
            self.worker.stop()
            self.worker.wait(3000)
            self.worker = None

        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.source_list.setEnabled(True)
        self.dest_combo.setEnabled(True)

    def _on_status_changed(self, status: str):
        """Handle status changes."""
        if status == "running":
            self.status_label.setText("Status: ● Running")
            self.status_label.setStyleSheet(
                "font-weight: bold; color: #4CAF50; padding: 5px;"
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
    app.setApplicationName(APP_NAME)

    # Set app icon
    app.setWindowIcon(get_app_icon())

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
