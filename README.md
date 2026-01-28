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
- [Git for Windows](https://git-scm.com/download/win)
- Python 3.9+ (uv installs it automatically)
- Administrator privileges (required for spool file access)

## Installation

### Option 1: Using uv (Recommended)

The easiest way to install and run using [uv](https://docs.astral.sh/uv/):

```powershell
# Install uv
irm https://astral.sh/uv/install.ps1 | iex

# Clone and run
git clone https://github.com/EmiliaVision/emilia-print-mirror.git
cd emilia-print-mirror
uv sync
uv run emilia-mirror
```

### Option 2: Run from Source (pip)

```powershell
# Clone the repository
git clone https://github.com/EmiliaVision/emilia-print-mirror.git
cd emilia-print-mirror

# Create virtual environment and install dependencies
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# Run the GUI application
python src/mirror_app.py
```

### Option 3: Install as Windows Service (with NSSM)

For running as a background service, use Python directly with [NSSM](https://nssm.cc/).

```powershell
# 1. Clone the repository
git clone https://github.com/EmiliaVision/emilia-print-mirror.git
cd emilia-print-mirror

# 2. Create virtual environment
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 3. Download NSSM from https://nssm.cc/download and extract to C:\Tools\nssm\

# 4. Create configuration directory and file
New-Item -ItemType Directory -Force -Path "C:\ProgramData\EmiliaPrintMirror"

# 5. Create config.json with your printers
@"
{
  "source_printers": ["YourSourcePrinter1", "YourSourcePrinter2"],
  "dest_printer": "YourDestinationPrinter",
  "interval": 1.0
}
"@ | Out-File -FilePath "C:\ProgramData\EmiliaPrintMirror\config.json" -Encoding UTF8

# 6. Install the service with NSSM (adjust paths as needed)
$repoPath = "C:\Users\$env:USERNAME\emilia-print-mirror"
nssm install EmiliaPrintMirror "$repoPath\venv\Scripts\python.exe" "$repoPath\src\mirror_service.py console"

# 7. Configure NSSM logging
nssm set EmiliaPrintMirror AppStdout "C:\ProgramData\EmiliaPrintMirror\service_stdout.log"
nssm set EmiliaPrintMirror AppStderr "C:\ProgramData\EmiliaPrintMirror\service_stderr.log"

# 8. Configure service to run as SYSTEM (for spool access)
nssm set EmiliaPrintMirror ObjectName LocalSystem

# 9. Start the service
nssm start EmiliaPrintMirror
```

**Service Management Commands:**

```powershell
# Check status
nssm status EmiliaPrintMirror
Get-Service EmiliaPrintMirror

# View logs
Get-Content "C:\ProgramData\EmiliaPrintMirror\service_stderr.log" -Tail 50

# Restart service
nssm restart EmiliaPrintMirror

# Stop service
nssm stop EmiliaPrintMirror

# Remove service
nssm remove EmiliaPrintMirror confirm
```

### Option 4: Build GUI Executable from Source

> **Note:** The .exe build currently has issues with win32print. Use Options 1-3 instead.

```powershell
# Install dependencies
pip install -r requirements.txt

# Build GUI application
pyinstaller build_mirror.spec
# Output: dist/EmiliaPrintMirror.exe
```

## Usage

### GUI Application

1. Run the application as Administrator (right-click > Run as administrator)
2. Select one or more **SOURCE** printers (Ctrl+click for multiple)
3. Select the **DESTINATION** printer
4. Click **Start Mirror**
5. Print to any source printer - it will be automatically copied to the destination
6. Check "Auto-start mirror when application opens" to start automatically next time

#### Auto-start GUI on Windows Login

To run the GUI automatically when you log in (as Administrator), create a `.bat` file and schedule it:

```powershell
# 1. Create a batch file to launch the app
$repoPath = "C:\Users\$env:USERNAME\emilia-print-mirror"
@"
@echo off
cd /d $repoPath
call venv\Scripts\python.exe src\mirror_app.py
"@ | Out-File -FilePath "$repoPath\EmiliaPrintMirror.bat" -Encoding ASCII

# 2. Create scheduled task that runs at login with admin privileges
$Action = New-ScheduledTaskAction -Execute "$repoPath\EmiliaPrintMirror.bat"
$Trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest

Register-ScheduledTask -TaskName "EmiliaPrintMirror" -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Force
```

To remove the auto-start:
```powershell
Unregister-ScheduledTask -TaskName "EmiliaPrintMirror" -Confirm:$false
```

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
1. Check log file at `C:\ProgramData\EmiliaPrintMirror\service_stderr.log`
2. Run in console mode to see errors:
   ```powershell
   python src/mirror_service.py console
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
