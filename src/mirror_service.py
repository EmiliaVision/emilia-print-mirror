"""
Spooler Queue Copy - Windows Service

Windows service for automatic printer mirroring.
Can be installed as a system service.
"""

import sys
import os
import time
import json
import logging
from typing import Set, Optional
from pathlib import Path

# Default configuration
DEFAULT_CONFIG = {
    "source_printer": "EmiliaCloudPrinterEpsonOrg",
    "dest_printer": "EmiliaCloudPrinterEpsonCopy",
    "interval": 1.0,
    "log_file": "mirror_service.log",
}

CONFIG_PATH = (
    Path(os.environ.get("PROGRAMDATA", "C:\\ProgramData"))
    / "SpoolerQueueCopy"
    / "config.json"
)


def get_config() -> dict:
    """Read configuration from file."""
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r") as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
        except:
            pass
    return DEFAULT_CONFIG


def save_config(config: dict):
    """Save configuration to file."""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


try:
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    import win32print

    SERVICE_AVAILABLE = True
except ImportError:
    SERVICE_AVAILABLE = False


class PrinterMirrorCore:
    """Core mirror service (no GUI dependencies)."""

    def __init__(
        self, source_printer: str, dest_printer: str, interval: float = 1.0, logger=None
    ):
        self.source_printer = source_printer
        self.dest_printer = dest_printer
        self.interval = interval
        self.logger = logger or logging.getLogger(__name__)
        self.running = False
        self.processed_jobs: Set[int] = set()
        self.spool_dir = os.path.join(
            os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS"
        )

    def log(self, message: str):
        self.logger.info(message)

    def _find_spool_file(self, job_id: int) -> Optional[str]:
        """Find the SPL file for a specific job."""
        try:
            patterns = [f"FP{job_id:05d}.SPL", f"{job_id:05d}.SPL"]
            for pattern in patterns:
                filepath = os.path.join(self.spool_dir, pattern)
                if os.path.exists(filepath):
                    return filepath

            # Find most recent SPL file
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
            self.log(f"Error finding spool file: {e}")
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
                    f"OK: Job {job_id} -> {self.dest_printer} (new: {new_job_id}, {len(spool_data)} bytes)"
                )
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
                    jobs[job_id] = {
                        "document": job.get("pDocument", "Unknown"),
                        "status": job.get("Status", 0),
                    }
            finally:
                win32print.ClosePrinter(handle)
        except Exception:
            pass
        return jobs

    def run_once(self) -> int:
        """Execute one iteration of the mirror. Returns number of jobs copied."""
        copied = 0
        current_jobs = self._get_current_jobs()
        new_job_ids = set(current_jobs.keys()) - self.processed_jobs

        for job_id in sorted(new_job_ids):
            if not self.running:
                break

            job_info = current_jobs[job_id]
            document = job_info["document"]

            # Skip already mirrored jobs
            if document.startswith("[MIRROR]") or document.startswith("[COPY]"):
                self.processed_jobs.add(job_id)
                continue

            self.log(f"New job detected: [{job_id}] {document}")
            time.sleep(1.0)

            if self._copy_job(job_id, document):
                copied += 1
            self.processed_jobs.add(job_id)

        # Clean up processed jobs that no longer exist
        self.processed_jobs &= set(current_jobs.keys())
        return copied

    def run(self):
        """Run the main mirror loop."""
        self.running = True
        self.log(f"Mirror started: {self.source_printer} -> {self.dest_printer}")

        try:
            os.listdir(self.spool_dir)
        except PermissionError:
            self.log(
                "ERROR: No access to spool directory. Run as Administrator/SYSTEM."
            )
            return

        existing_jobs = self._get_current_jobs()
        self.processed_jobs = set(existing_jobs.keys())
        self.log(f"Ignoring {len(self.processed_jobs)} existing job(s)")

        while self.running:
            self.run_once()
            time.sleep(self.interval)

        self.log("Mirror stopped")

    def stop(self):
        """Stop the mirror service."""
        self.running = False


if SERVICE_AVAILABLE:

    class SpoolerQueueCopyService(win32serviceutil.ServiceFramework):
        """Windows Service for Spooler Queue Copy."""

        _svc_name_ = "SpoolerQueueCopy"
        _svc_display_name_ = "Spooler Queue Copy Mirror Service"
        _svc_description_ = (
            "Automatically copies print jobs from one printer to another"
        )

        def __init__(self, args):
            win32serviceutil.ServiceFramework.__init__(self, args)
            self.stop_event = win32event.CreateEvent(None, 0, 0, None)
            self.mirror = None

            # Configure logging
            log_path = (
                Path(os.environ.get("PROGRAMDATA", "C:\\ProgramData"))
                / "SpoolerQueueCopy"
                / "service.log"
            )
            log_path.parent.mkdir(parents=True, exist_ok=True)

            logging.basicConfig(
                level=logging.INFO,
                format="%(asctime)s - %(levelname)s - %(message)s",
                handlers=[logging.FileHandler(str(log_path)), logging.StreamHandler()],
            )
            self.logger = logging.getLogger(__name__)

        def SvcStop(self):
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self.stop_event)
            if self.mirror:
                self.mirror.stop()

        def SvcDoRun(self):
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, ""),
            )
            self.main()

        def main(self):
            config = get_config()

            self.mirror = PrinterMirrorCore(
                source_printer=config["source_printer"],
                dest_printer=config["dest_printer"],
                interval=config["interval"],
                logger=self.logger,
            )

            self.mirror.running = True
            self.logger.info(
                f"Service started: {config['source_printer']} -> {config['dest_printer']}"
            )

            # Initialize
            try:
                os.listdir(self.mirror.spool_dir)
            except PermissionError:
                self.logger.error("No access to spool directory")
                return

            existing = self.mirror._get_current_jobs()
            self.mirror.processed_jobs = set(existing.keys())

            # Main loop
            while self.mirror.running:
                rc = win32event.WaitForSingleObject(
                    self.stop_event, int(config["interval"] * 1000)
                )
                if rc == win32event.WAIT_OBJECT_0:
                    break
                self.mirror.run_once()

            self.logger.info("Service stopped")


def install_service():
    """Install the Windows service."""
    if not SERVICE_AVAILABLE:
        print("Error: pywin32 is not available")
        return False

    try:
        # Save default configuration first
        config = get_config()
        save_config(config)
        print(f"Configuration saved to: {CONFIG_PATH}")

        # Install service
        win32serviceutil.InstallService(
            SpoolerQueueCopyService._svc_name_,
            SpoolerQueueCopyService._svc_name_,
            SpoolerQueueCopyService._svc_display_name_,
            startType=win32service.SERVICE_AUTO_START,
            description=SpoolerQueueCopyService._svc_description_,
        )
        print(
            f"Service '{SpoolerQueueCopyService._svc_display_name_}' installed successfully"
        )
        return True
    except Exception as e:
        print(f"Error installing service: {e}")
        return False


def uninstall_service():
    """Uninstall the Windows service."""
    if not SERVICE_AVAILABLE:
        print("Error: pywin32 is not available")
        return False

    try:
        win32serviceutil.RemoveService(SpoolerQueueCopyService._svc_name_)
        print("Service uninstalled successfully")
        return True
    except Exception as e:
        print(f"Error uninstalling service: {e}")
        return False


def run_console():
    """Run in console mode (not as service)."""
    config = get_config()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )

    print(f"""
    ╔═══════════════════════════════════════════════════════════╗
    ║          SPOOLER QUEUE COPY - MIRROR SERVICE              ║
    ╚═══════════════════════════════════════════════════════════╝
    
    Source:      {config["source_printer"]}
    Destination: {config["dest_printer"]}
    
    Press Ctrl+C to stop
    """)

    mirror = PrinterMirrorCore(
        source_printer=config["source_printer"],
        dest_printer=config["dest_printer"],
        interval=config["interval"],
    )

    try:
        mirror.run()
    except KeyboardInterrupt:
        mirror.stop()
        print("\nStopped by user")


def main():
    """Main entry point."""
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()

        if cmd == "install":
            install_service()
        elif cmd == "uninstall" or cmd == "remove":
            uninstall_service()
        elif cmd == "start":
            os.system(f"net start {SpoolerQueueCopyService._svc_name_}")
        elif cmd == "stop":
            os.system(f"net stop {SpoolerQueueCopyService._svc_name_}")
        elif cmd == "restart":
            os.system(f"net stop {SpoolerQueueCopyService._svc_name_}")
            time.sleep(2)
            os.system(f"net start {SpoolerQueueCopyService._svc_name_}")
        elif cmd == "console":
            run_console()
        elif cmd == "config":
            if len(sys.argv) >= 4:
                config = get_config()
                config["source_printer"] = sys.argv[2]
                config["dest_printer"] = sys.argv[3]
                save_config(config)
                print(f"Configuration updated:")
                print(f"  Source:      {config['source_printer']}")
                print(f"  Destination: {config['dest_printer']}")
            else:
                print("Usage: mirror_service.py config <source> <destination>")
        elif cmd == "status":
            config = get_config()
            print(f"Current configuration ({CONFIG_PATH}):")
            print(f"  Source:      {config['source_printer']}")
            print(f"  Destination: {config['dest_printer']}")
            print(f"  Interval:    {config['interval']}s")
        else:
            print(f"""
Spooler Queue Copy - Mirror Service

Usage: {sys.argv[0]} <command>

Commands:
  install     - Install Windows service
  uninstall   - Uninstall Windows service
  start       - Start the service
  stop        - Stop the service
  restart     - Restart the service
  console     - Run in console mode (not as service)
  config <source> <dest> - Configure printers
  status      - Show current configuration
            """)
    else:
        # If run without arguments, try as service
        if SERVICE_AVAILABLE and len(sys.argv) == 1:
            try:
                servicemanager.Initialize()
                servicemanager.PrepareToHostSingle(SpoolerQueueCopyService)
                servicemanager.StartServiceCtrlDispatcher()
            except Exception:
                # If it fails, run in console mode
                run_console()
        else:
            run_console()


if __name__ == "__main__":
    main()
