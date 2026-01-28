<div align="center">
  <img src="assets/banner.png" alt="Emilia Print Mirror Banner" width="100%">
</div>

# Emilia Print Mirror

A Windows application to automatically mirror print jobs from one or more printers to a destination printer.

## Features

- **Multi-Source Support**: Monitor multiple source printers simultaneously
- **Automatic Mirroring**: Jobs are automatically copied to the destination printer
- **GUI Application**: Easy-to-use graphical interface
- **Windows Service**: Run as a background service that starts with Windows
- **One-Click Install**: Install service directly from the GUI

## Quick Start

1. Download `EmiliaPrintMirror.exe` and `EmiliaMirrorService.exe` from [Releases](https://github.com/EmiliaVision/emilia-print-mirror/releases/latest)
2. Run `EmiliaPrintMirror.exe` as Administrator
3. Select source printer(s) and destination printer
4. Click **"Install Service"**

For a detailed visual guide, see **[Installation Guide](docs/INSTALL_GUIDE.md)**.

## Requirements

- Windows 10 or Windows 11
- Administrator privileges

## How It Works

<div align="center">
  <img src="assets/how-it-works.png" alt="Emilia Print Mirror Flow" width="100%">
</div>

The service monitors the Windows print spooler. When a new job is detected on a source printer, it reads the spool file and sends the raw print data to the destination printer.

## Configuration

Configuration is stored at:
- **Service config**: `C:\ProgramData\EmiliaPrintMirror\config.json`
- **GUI config**: `%APPDATA%\EmiliaPrintMirror\gui_config.json`

Example config:
```json
{
  "source_printers": ["Printer1", "Printer2"],
  "dest_printer": "DestinationPrinter",
  "interval": 1.0
}
```

## Logs

Service logs are stored at:
```
C:\ProgramData\EmiliaPrintMirror\service.log
```

## Alternative Installation Methods

### Using uv (for development)

```powershell
git clone https://github.com/EmiliaVision/emilia-print-mirror.git
cd emilia-print-mirror
uv sync
uv run emilia-mirror        # Run GUI
uv run emilia-mirror-service console  # Run service in console mode
```

### Build from Source

```powershell
uv sync --all-extras
uv run pyinstaller build_mirror.spec   # Build GUI
uv run pyinstaller build_service.spec  # Build Service
```

## Troubleshooting

### "Access Denied" when installing service
Run the GUI as Administrator (right-click → Run as administrator).

### Service won't start
Check logs at `C:\ProgramData\EmiliaPrintMirror\service.log`.

### Jobs not being copied
Ensure **KeepPrintedJobs** is enabled on source printers:
```powershell
Set-Printer -Name "YourSourcePrinter" -KeepPrintedJobs $true
```

## Project Structure

```
emilia-print-mirror/
├── src/
│   ├── mirror_app.py      # GUI application (PyQt6)
│   └── mirror_service.py  # Windows service
├── assets/
│   └── screenshots/       # Installation guide images
├── docs/
│   └── INSTALL_GUIDE.md   # Visual installation guide
├── build_mirror.spec      # PyInstaller config (GUI)
├── build_service.spec     # PyInstaller config (Service)
└── README.md
```

## License

MIT License

## Credits

- Icon: Flower icon from [Phosphor Icons](https://phosphoricons.com/)
