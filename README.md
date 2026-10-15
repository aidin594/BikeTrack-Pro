<div align="center">

![BikeTrack Pro banner](assets/biketrack-banner.svg)

# BikeTrack Pro

**Bicycle workshop management for Windows**

A local-first desktop application for repair tracking, customer visit history and receipt printing.

[![Windows](https://img.shields.io/badge/Platform-Windows-0078D4)](https://github.com/aidin594/BikeTrack-Pro)
[![Python](https://img.shields.io/badge/Python-3.11-3776AB)](https://www.python.org)
[![Windows build](https://github.com/aidin594/BikeTrack-Pro/actions/workflows/build-windows.yml/badge.svg)](https://github.com/aidin594/BikeTrack-Pro/actions/workflows/build-windows.yml)

[**Download**](https://github.com/aidin594/BikeTrack-Pro/releases) · [**Issues**](https://github.com/aidin594/BikeTrack-Pro/issues) · [**Builds**](https://github.com/aidin594/BikeTrack-Pro/actions)

</div>

## What is BikeTrack Pro?

BikeTrack Pro is a Python/Tkinter application developed for the daily operations of a bicycle repair workshop. Records are stored locally using SQLite; no cloud account is required.

## Features

| Feature | Description |
| --- | --- |
| Bicycle repair tracking | Register bicycle models and up to 10 repair items per visit |
| Visit history | Find previous bicycle visits by phone number |
| Repair editing | Add, change, reorder or remove repair items |
| Shop-only notes | Print internal instructions only on the workshop receipt |
| Receipt settings | Configure customer and shop printer settings separately |
| Customer flags | Highlight customer numbers requiring special attention |
| Bike statistics | Compare bicycle record counts across years |
| SQLite backups | Configure a database backup folder |
| Built-in Help | Offline help with tabs for common workflows |

## Application screenshots

The header is original project artwork, **not a screenshot**. Actual Windows screenshots of the main screen, visit history, receipt settings and Help will be added after capture and verification.

## Download the Windows application

Go to [**Releases**](https://github.com/aidin594/BikeTrack-Pro/releases) and download `BikeTrack-Pro.exe` from the latest available release. This is a portable executable, not a Windows installer.

If the latest release is unavailable, download the `BikeTrack-Pro-Windows` ZIP artifact from a successful [GitHub Actions build](https://github.com/aidin594/BikeTrack-Pro/actions/workflows/build-windows.yml) (GitHub sign-in may be required).

**Important:** Before replacing an older version, back up the existing `customer_bikes.db`, settings and receipt configuration. Test the new executable with sample records before production use.

## Run from source

Requires Windows 10/11 and Python 3.11 (recommended). Printer functions require compatible Windows drivers.

```powershell
git clone https://github.com/aidin594/BikeTrack-Pro.git
cd BikeTrack-Pro
py -m pip install -r requirements.txt
py BikeTrack_Pro.pyw
```

Dependencies include Pillow, qrcode, python-Levenshtein and pywin32.

## Automatic builds

A Windows EXE is built automatically with PyInstaller and GitHub Actions whenever new code is pushed to `main`. Successful compilation does not establish compatibility with every printer or guarantee all workflows have been tested.

## Privacy and safety

The app stores customer details and repair records in a local SQLite database. **Never upload real databases, personal information, local settings, backups or customer exports to GitHub.** The `.gitignore` helps prevent accidental commits but is not a substitute for checking changes before publishing.

## Support

[Report a bug or suggest a feature](https://github.com/aidin594/BikeTrack-Pro/issues). Do not include customer personal data in public issues.

Developed by [Aidin594](https://github.com/aidin594) for bicycle workshop operations. Contact: info@vhcyklar.se.

## License

No open-source license has been selected. Source availability does not automatically authorize modification, redistribution or commercial use.
