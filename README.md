# OpenText OTDS Bulk Provisioning Tool

[![Build](https://img.shields.io/github/actions/workflow/status/YOUR_USERNAME/otds-bulk-tool/build.yml?label=Build&logo=github&style=flat-square)](https://github.com/YOUR_USERNAME/otds-bulk-tool/actions)
[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?style=flat-square&logo=python)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey?style=flat-square)](https://github.com/YOUR_USERNAME/otds-bulk-tool/releases)
[![OTDS](https://img.shields.io/badge/OpenText%20OTDS-25.x-0063a5?style=flat-square)](https://developer.opentext.com/ce/products/opentext-directory-services)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)

> Bulk create users and groups in OpenText Directory Services (OTDS) 25.x via REST API,
> organized per partition, with dry-run mode, retry logic, and HTML reporting.

## Features
- OTDS 25.x REST API using /otdsws/rest with OTDSTicket authentication
- Partition-aware: each Excel sheet = one OTDS partition
- Bulk Users: login, name, email, password, title, department, company, phone
- Bulk Groups: name, description, email, member list
- Auto-membership via semicolon-separated groups column
- 3-attempt retry logic with 2s delay
- Dry-run mode: preview without touching OTDS
- Color-coded HTML report: Created / Skipped / Errors
- Console + file logging with timestamps
- HTTP 409 = already exists -> logged as Skipped, not Error
- Cross-platform: Windows .exe, Linux, macOS via GitHub Actions

## Download Pre-built Executables

| Platform | Download |
|----------|----------|
| Windows  | otds_bulk_tool-windows.exe |
| Linux    | otds_bulk_tool-linux       |
| macOS    | otds_bulk_tool-macos       |

Linux/macOS: `chmod +x otds_bulk_tool-linux && ./otds_bulk_tool-linux --init`

## Quick Start

```bash
# Option A: Pre-built executable
./otds_bulk_tool-linux --init         # Generate templates
# Edit config.json
./otds_bulk_tool-linux --dry-run      # Preview
./otds_bulk_tool-linux                # Run live

# Option B: Python source
pip install -r requirements.txt
python otds_bulk_tool.py --init
python otds_bulk_tool.py --dry-run
python otds_bulk_tool.py --otds-url https://otds.yourserver.com:8443
```

## CLI Reference

| Argument        | Description                                  | Default       |
|-----------------|----------------------------------------------|---------------|
| --init          | Generate config.json + Excel templates       | -             |
| --config FILE   | Config JSON file path                        | config.json   |
| --otds-url URL  | Override OTDS base URL                       | from config   |
| --users XLSX    | Override users Excel path                    | from config   |
| --groups XLSX   | Override groups Excel path                   | from config   |
| --dry-run       | Preview without API calls                    | false         |
| --log-level     | DEBUG/INFO/WARNING/ERROR                     | INFO          |
| --report HTML   | Override HTML report path                    | from config   |

## Excel Format

### users.xlsx (one sheet per OTDS partition)
Columns: login, first_name, last_name, email, password, title, department, company, phone, description, groups
- groups column: semicolon-separated group names e.g. ECM_Admins;CS_Users

### groups.xlsx (one sheet per OTDS partition)
Columns: name, description, email, members
- members column: semicolon-separated logins e.g. naveen.k;vijo.j

## config.json Keys

| Key          | Description                  | Default                         |
|--------------|------------------------------|---------------------------------|
| otds_url     | OTDS base URL with port      | https://otds.example.com:8443   |
| admin_user   | Admin account                | otadmin@otds.admin              |
| admin_pass   | Admin password               | password                        |
| verify_ssl   | Verify SSL certificate       | false                           |
| timeout      | HTTP timeout seconds         | 30                              |
| dry_run      | Dry-run mode                 | false                           |
| log_level    | Logging verbosity            | INFO                            |
| log_file     | Log output file              | otds_bulk.log                   |
| report_file  | HTML report output           | otds_bulk_report.html           |
| users_input  | Users Excel file             | users.xlsx                      |
| groups_input | Groups Excel file            | groups.xlsx                     |

## GitHub Actions – Build Your Own Executable

```bash
# Tag and push to trigger build
git tag v1.0.0
git push origin v1.0.0
# Or: Actions tab -> Run workflow -> Enter version
```

## Build Locally

```bash
# Linux/macOS
chmod +x build_local.sh && ./build_local.sh

# Windows
build_local.bat
```

## Project Structure

```
otds-bulk-tool/
├── .github/workflows/build.yml   GitHub Actions CI/CD
├── otds_bulk_tool.py             Main Python script
├── otds_bulk_tool.spec           PyInstaller build spec
├── requirements.txt              Python dependencies
├── build_local.sh                Local build (Linux/macOS)
├── build_local.bat               Local build (Windows)
├── setup.py                      Python package setup
└── README.md                     This file
```

## License

MIT License - Copyright (c) 2026 Naveen K. Manam

Built for OpenText ECM/BPM environments. Compatible with OTDS 22.x / 24.x / 25.x
