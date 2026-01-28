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
- Administrator privileges (required for spool file access)
- Python 3.9+ (only for development)

## Installation

### Option 1: Pre-built Executables (Recommended)

Download the pre-built executables from the `dist/` folder:

| File | Description |
|------|-------------|
| `EmiliaPrintMirror.exe` | GUI application (35 MB) |
| `EmiliaPrintMirrorService.exe` | Background service (7 MB) |

#### Install as GUI Application

1. Copy `EmiliaPrintMirror.exe` to your preferred location (e.g., `C:\Program Files\EmiliaPrintMirror\`)
2. Create a desktop shortcut (optional)
3. Right-click the shortcut > Properties > Advanced > **Run as administrator**
4. Double-click to run

#### Install as Windows Service (with NSSM)

[NSSM](https://nssm.cc/) (Non-Sucking Service Manager) is recommended for running the service.

```powershell
# 1. Download NSSM from https://nssm.cc/download
# 2. Extract and add to PATH, or use full path to nssm.exe

# 3. Create configuration directory and file
New-Item -ItemType Directory -Force -Path "C:\ProgramData\EmiliaPrintMirror"

# 4. Create config.json with your printers
@"
{
  "source_printers": ["YourSourcePrinter1", "YourSourcePrinter2"],
  "dest_printer": "YourDestinationPrinter",
  "interval": 1.0
}
"@ | Out-File -FilePath "C:\ProgramData\EmiliaPrintMirror\config.json" -Encoding UTF8

# 5. Install the service with NSSM
nssm install EmiliaPrintMirror "C:\Program Files\EmiliaPrintMirror\EmiliaPrintMirrorService.exe" console

# 6. Configure service to run as SYSTEM (for spool access)
nssm set EmiliaPrintMirror ObjectName LocalSystem

# 7. Start the service
nssm start EmiliaPrintMirror
```

**Service Management Commands:**

```powershell
# Check status
nssm status EmiliaPrintMirror

# View logs
Get-Content "C:\ProgramData\EmiliaPrintMirror\service.log" -Tail 50

# Restart service
nssm restart EmiliaPrintMirror

# Stop service
nssm stop EmiliaPrintMirror

# Remove service
nssm remove EmiliaPrintMirror confirm
```

### Option 2: Run from Source

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

### Option 3: Using uv

```powershell
# Install uv
irm https://astral.sh/uv/install.ps1 | iex

# Install dependencies and run
uv sync
uv run python src/mirror_app.py
```

### Option 4: Build Executables from Source

```powershell
# Install dependencies
pip install -r requirements.txt

# Build GUI application
pyinstaller build_mirror.spec
# Output: dist/EmiliaPrintMirror.exe

# Build Service executable
pyinstaller --onefile --noconsole --name EmiliaPrintMirrorService src/mirror_service.py
# Output: dist/EmiliaPrintMirrorService.exe
```

## Usage

### GUI Application

1. Run `EmiliaPrintMirror.exe` as Administrator
2. Select one or more **SOURCE** printers (Ctrl+click for multiple)
3. Select the **DESTINATION** printer
4. Click **Start Mirror**
5. Print to any source printer - it will be automatically copied to the destination
6. Minimize to system tray (optional)

### Windows Service (Command Line)

If running from source:

```powershell
# Configure printers (run once)
python src/mirror_service.py config "SourcePrinter1,SourcePrinter2" "DestinationPrinter"

# Run in console mode (for testing)
python src/mirror_service.py console

# Check current configuration
python src/mirror_service.py status
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
- Run the application as Administrator
- If running as service, ensure it runs as `LocalSystem`

### "Could not read job"
1. Ensure `KeepPrintedJobs` is enabled on source printer(s):
   ```powershell
   Set-Printer -Name "YourSourcePrinter" -KeepPrintedJobs $true
   ```
2. Try increasing the interval between checks

### Service won't start
1. Check log file at `C:\ProgramData\EmiliaPrintMirror\service.log`
2. Run in console mode to see errors:
   ```powershell
   EmiliaPrintMirrorService.exe console
   ```
3. Verify config.json exists and has valid JSON

### GUI won't start
1. Ensure you're running as Administrator
2. Check if PyQt6 dependencies are available (if running from source)
3. Try running from command line to see error messages

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
