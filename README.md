# Office Setup Wizard

<div align="center">

[![Download OfficeSetupWizard.exe](https://img.shields.io/badge/Download-OfficeSetupWizard.exe-2ea44f?style=for-the-badge&logo=windows11&logoColor=white)](https://github.com/TARIKTR1099/OfficeSetupWizard/releases/latest/download/OfficeSetupWizard.exe)

[![Release](https://img.shields.io/github/v/release/TARIKTR1099/OfficeSetupWizard)](https://github.com/TARIKTR1099/OfficeSetupWizard/releases/latest)

</div>

A Windows GUI tool that installs Microsoft Office **LTSC Professional Plus 2024**
with a single click. It wraps the official Microsoft Office Deployment Tool (ODT)
and adds an easy wizard for picking your language, apps and configuration.

Built with Python + Tkinter and packaged as a standalone `.exe` with PyInstaller.

## Features

- **One-click install** — downloads and installs Office LTSC Professional Plus 2024
- **Language picker** — type to filter; matching languages move to the top of the
  dropdown while the full list stays visible below
- **Custom configuration** — paste your own XML from
  [config.office.com](https://config.office.com/deploymentsettings)
- **Built-in editor** — tweak edition (64-bit), channel, apps, Visio/Project,
  auto-activate and more before installing
- **Smart reinstall** — detects an existing Office install and offers
  uninstall & reinstall or activate-only
- **Turkish & English UI**, light/dark/auto theme
- **Activates Office (Ohook)** via Microsoft Activation Scripts after setup

## Requirements

- Windows 10/11 (64-bit)
- Internet connection (Office files are downloaded from the Microsoft CDN)
- Administrator rights (requested automatically via UAC)

## Usage

Run `OfficeSetupWizard.exe` as administrator:

1. Pick a setup mode
   - **Automatic Setup** — choose your language, optionally edit advanced settings
   - **Custom Configuration** — load your own XML file from config.office.com
2. Click **Start Setup** and wait. Office is downloaded, installed and activated.

## Building from source

```
git clone https://github.com/TARIKTR1099/OfficeSetupWizard.git
cd OfficeSetupWizard
powershell -ExecutionPolicy Bypass -File builder\build.ps1
```

The build script:

- installs `pyinstaller`
- bundles the Office Deployment Tool files from `resources/`
- outputs `OfficeSetupWizard.exe` to the repository root

The `resources/` folder must contain the ODT files (`setup.exe` and
`officedeploymenttool_20131-20090.exe`) downloaded from [aka.ms/odt](https://aka.ms/odt).

## Project layout

```
OfficeSetupWizard/
├── builder/
│   ├── office_setup_wizard.py    # main source
│   ├── office_setup_wizard.spec  # PyInstaller spec (relative paths)
│   ├── office_setup_wizard.ico   # application icon
│   ├── build.ps1                 # one-command build script
│   ├── version_info.txt          # EXE version metadata
│   └── icons/                    # theme icons (light/dark)
├── resources/
│   ├── setup.exe                 # Office Deployment Tool
│   ├── officedeploymenttool_20131-20090.exe
│   ├── Configuration.xml         # default install configuration
│   └── configuration-Office365-x64.xml
└── OfficeSetupWizard.exe         # build output (generated)
```

## License

This project uses the [Microsoft Office Deployment Tool](https://aka.ms/odt) and
[Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts).
You are responsible for complying with Microsoft licensing terms for Office and
for your own use of activation scripts. Distributing Office binaries may be
restricted; please review the Microsoft Software License Terms.
