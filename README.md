# Emilia Print Mirror

A Windows application to automatically mirror print jobs from one or more printers to a destination printer.

## Features

- **Multi-Source Support**: Monitor multiple source printers simultaneously
- **Automatic Print Mirroring**: Jobs are automatically copied to the destination printer
- **GUI Application**: Easy-to-use graphical interface with the Emilia flower logo
- **Windows Service**: Can run as a background Windows service
- **Real-time Monitoring**: Monitors print queues in real-time

## Requirements

- Windows 10 or Windows 11
- Python 3.9+ (for development)
- Administrator privileges (recommended for spool file access)

## Quick Start

### Option 1: Run from Source

```powershell
# Clone the repository
git clone <repository-url>
cd emilia-print-mirror

# Create virtual environment and install dependencies
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# Run the GUI application
python src/mirror_app.py
```

### Option 2: Using uv (Recommended)

```powershell
# Install uv
irm https://astral.sh/uv/install.ps1 | iex

# Install dependencies and run
uv sync
uv run python src/mirror_app.py
```

### Option 3: Build Executable

```powershell
pip install -r requirements.txt
pyinstaller build_mirror.spec
# Executable will be in dist/EmiliaPrintMirror.exe
```

## Usage

### GUI Application

1. Run the application as Administrator
2. Select one or more **SOURCE** printers (Ctrl+click for multiple)
3. Select the **DESTINATION** printer
4. Click **Start Mirror**
5. Print to any source printer - it will be automatically copied to the destination

### Windows Service

```powershell
# Configure printers (run once)
python src/mirror_service.py config "SourcePrinter1,SourcePrinter2" "DestinationPrinter"

# Install as Windows service
python src/mirror_service.py install

# Start the service
python src/mirror_service.py start

# Check status
python src/mirror_service.py status

# Stop the service
python src/mirror_service.py stop

# Uninstall the service
python src/mirror_service.py uninstall

# Run in console mode (for testing)
python src/mirror_service.py console
```

## Configuration

### Printer Requirements

The source printer(s) must have **KeepPrintedJobs** enabled:

```powershell
Set-Printer -Name "YourSourcePrinter" -KeepPrintedJobs $true
```

The GUI application can do this automatically if you check the "Auto-configure printers" option.

### Configuration File

Service configuration is stored in:
```
C:\ProgramData\EmiliaPrintMirror\config.json
```

Example:
```json
{
  "source_printers": ["EmiliaCloudPrinterEpsonOrg", "EmiliaCloudPrinterEpsonOrg2"],
  "dest_printer": "EmiliaCloudPrinterEpsonCopy",
  "interval": 1.0
}
```

## How It Works

```
┌─────────────────────┐
│ Source Printer 1    │──┐
└─────────────────────┘  │
                         │    ┌─────────────────────┐
┌─────────────────────┐  ├───▶│ Destination Printer │
│ Source Printer 2    │──┤    └─────────────────────┘
└─────────────────────┘  │
                         │
┌─────────────────────┐  │
│ Source Printer N    │──┘
└─────────────────────┘
```

1. The application monitors Windows print spooler for new jobs on source printer(s)
2. When a new job is detected, it reads the spool file from `C:\Windows\System32\spool\PRINTERS`
3. The raw print data is sent to the destination printer
4. Jobs prefixed with `[MIRROR]` are ignored to prevent infinite loops

## Creating Test Printers

```powershell
# Create TCP/IP printer port
Add-PrinterPort -Name "EmiliaCloud-Test" -PrinterHostAddress "printer.example.com" -PortNumber 9100

# Create printers with Epson driver
Add-Printer -Name "EmiliaPrinterOrg" -DriverName "EPSON TM-T20II Receipt5" -PortName "EmiliaCloud-Test"
Add-Printer -Name "EmiliaPrinterCopy" -DriverName "EPSON TM-T20II Receipt5" -PortName "EmiliaCloud-Test"

# Enable KeepPrintedJobs on source
Set-Printer -Name "EmiliaPrinterOrg" -KeepPrintedJobs $true
```

## Troubleshooting

### "No access to spool directory"
Run the application as Administrator.

### "Could not read job"
1. Ensure `KeepPrintedJobs` is enabled on source printer(s)
2. Try increasing the interval between checks

### Service won't start
1. Check log file at `C:\ProgramData\EmiliaPrintMirror\service.log`
2. Run `python src/mirror_service.py console` to see errors

## Project Structure

```
emilia-print-mirror/
├── src/
│   ├── mirror_app.py      # GUI application (PyQt6)
│   ├── mirror_service.py  # Windows service / CLI
│   └── __init__.py
├── assets/
│   └── icon.svg           # Emilia flower icon
├── requirements.txt
├── build_mirror.spec      # PyInstaller config
├── pyproject.toml
└── README.md
```

## License

MIT License

## Credits

- Icon: Flower icon from [Phosphor Icons](https://phosphoricons.com/)
