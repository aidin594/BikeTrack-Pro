# BikeTrack Pro

Windows desktop application for bicycle repair workshops, built with Python and Tkinter.

## Features

- Customer and bicycle repair records with visit history
- Repair issue tracking and editing
- Shop-only internal notes on workshop receipts
- Separate customer and workshop receipt printer settings
- Customer flags and bike statistics
- SQLite database with backup configuration

## Requirements

- Windows 10/11
- Python 3.10+ (recommended)
- A configured receipt printer for printing receipts

## Run locally

```powershell
py -m pip install -r requirements.txt
py BikeTrack_Pro.pyw
```

The application stores customer records in a local SQLite database. **Do not upload real databases, customer data, local settings, backups or printer configuration to GitHub.** The included `.gitignore` excludes common sensitive/runtime files.

## Notes

- `icon.ico` is optional and not included in this source package.
- This is an initial source release, not a packaged Windows installer.
- Windows-only printing relies on `pywin32`.
- Source syntax has been checked; Windows GUI and printer behavior have not been tested as part of this package.

## License

No open-source license has been selected yet. Until a license is added, normal copyright restrictions apply.
