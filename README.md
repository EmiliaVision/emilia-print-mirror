# Spooler Queue Copy

A Windows application to automatically mirror print jobs from one printer to another.

## Features

- **Automatic Print Mirroring**: When you print to a source printer, the job is automatically copied to a destination printer
- **GUI Application**: Easy-to-use graphical interface for configuration
- **Windows Service**: Can run as a background Windows service
- **Real-time Monitoring**: Monitors print queues in real-time
- **Admin-free operation**: Works with proper configuration

## Requirements

- Windows 10 or Windows 11
- Python 3.9+ (for development)
- Administrator privileges (recommended for spool file access)

## Installation

### Option 1: Run from Source

```bash
# Clone the repository
git clone <repository-url>
cd spooler-queue-copy

# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the GUI application
python src/mirror_app.py
```

### Option 2: Using uv (Recommended)

```bash
# Install uv
irm https://astral.sh/uv/install.ps1 | iex

# Install dependencies
uv sync

# Run the application
uv run python src/mirror_app.py
```

### Option 3: Build Executable

```bash
# Install dependencies
pip install -r requirements.txt

# Build the executable
pyinstaller build_mirror.spec

# The executable will be in dist/SpoolerQueueCopy.exe
```

## Usage

### GUI Application

1. Run `SpoolerQueueCopy.exe` as Administrator
2. Select the **SOURCE** printer (where you will print)
3. Select the **DESTINATION** printer (where jobs will be copied)
4. Click **Start Mirror**
5. Print something to the source printer - it will be automatically copied to the destination

### Command Line Service

```bash
# Show help
python src/mirror_service.py help

# Configure printers
python src/mirror_service.py config "SourcePrinter" "DestinationPrinter"

# Run in console mode
python src/mirror_service.py console

# Install as Windows service
python src/mirror_service.py install

# Start the service
python src/mirror_service.py start

# Stop the service
python src/mirror_service.py stop

# Uninstall the service
python src/mirror_service.py uninstall
```

## Configuration

### Printer Setup

Before using the mirror, ensure the source printer has **KeepPrintedJobs** enabled:

```powershell
Set-Printer -Name "YourSourcePrinter" -KeepPrintedJobs $true
```

The GUI application can do this automatically if you check the "Auto-configure printer" option.

### Configuration File

The service stores its configuration in:
```
C:\ProgramData\SpoolerQueueCopy\config.json
```

Example configuration:
```json
{
  "source_printer": "EmiliaCloudPrinterEpsonOrg",
  "dest_printer": "EmiliaCloudPrinterEpsonCopy",
  "interval": 1.0
}
```

## How It Works

1. The application monitors the Windows print spooler for new jobs on the source printer
2. When a new job is detected, it reads the spool file from `C:\Windows\System32\spool\PRINTERS`
3. The raw print data is then sent to the destination printer
4. Jobs prefixed with `[MIRROR]` are ignored to prevent infinite loops

## Architecture

```
spooler-queue-copy/
├── src/
│   ├── mirror_app.py      # GUI application (PyQt6)
│   ├── mirror_service.py  # Windows service / CLI
│   ├── print_spooler.py   # Windows Print Spooler API wrapper
│   └── __init__.py
├── requirements.txt       # Python dependencies
├── build_mirror.spec      # PyInstaller configuration
├── pyproject.toml         # Project configuration (uv)
└── README.md
```

## Troubleshooting

### "No access to spool directory"

Run the application as Administrator. The spool files are in a protected system directory.

### "Could not read job"

1. Ensure `KeepPrintedJobs` is enabled on the source printer
2. Check that no other application is locking the spool files
3. Try increasing the interval between checks

### "Job not copied"

1. Verify both printers are online and accessible
2. Check the Windows Event Log for print spooler errors
3. Ensure the destination printer driver is compatible with the job data

### Service won't start

1. Check the log file at `C:\ProgramData\SpoolerQueueCopy\service.log`
2. Ensure the configuration file has valid printer names
3. Run `python src/mirror_service.py console` to see errors in real-time

## Creating Printers for Testing

```powershell
# Create a TCP/IP printer port
Add-PrinterPort -Name "TestPort" -PrinterHostAddress "printer.example.com" -PortNumber 9100

# Create printers with Epson driver
Add-Printer -Name "TestPrinterOrg" -DriverName "EPSON TM-T20II Receipt5" -PortName "TestPort"
Add-Printer -Name "TestPrinterCopy" -DriverName "EPSON TM-T20II Receipt5" -PortName "TestPort"

# Or with Generic driver
Add-Printer -Name "TestPrinterOrg" -DriverName "Generic / Text Only" -PortName "TestPort"
Add-Printer -Name "TestPrinterCopy" -DriverName "Generic / Text Only" -PortName "TestPort"
```

## License

MIT License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request
