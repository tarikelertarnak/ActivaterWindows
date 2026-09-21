# ActivaterWindows

> **One-click Microsoft Office LTSC Professional Plus 2024 installer, with Windows, Office and ESU activation managed from a single interface.**

`ActivaterWindows` is a user-friendly **portable Windows desktop application** that wraps the Office Deployment Tool (ODT) to automate installation and uses **Microsoft Activation Scripts (MAS v3.12)** for one-click activation. A single `.exe` — no installation required, zero leftovers when deleted.

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-1.0.0-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/UI-Tkinter-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
  <img src="https://img.shields.io/badge/ODT-16.0.20228-brightgreen?style=flat-square" />
  <img src="https://img.shields.io/badge/MAS-3.12-red?style=flat-square" />
</p>

---

## Features

| Feature | Description |
|---|---|
| **Portable** | Single `.exe`, no installation, runs on any machine |
| **Zero Leftover** | Temp files, logs, registry cleaned on exit |
| **One-Click Installation** | Office LTSC Professional Plus 2024 in your language |
| **Custom Configuration** | Use your own `Configuration.xml` file |
| **38+ Languages** | tr-tr, en-us, de-de, fr-fr and more |
| **Light/Dark Theme** | Automatic system theme detection |
| **Embedded Activation (MAS)** | Windows / Office / ESU activation scripts are embedded in the exe — nothing is downloaded at runtime |
| **Activation Type Selection** | Windows 11 Pro / for Workstations / 11 / 10 Pro / 10 + Office — 16 methods |
| **Activation Status Check** | One-click current activation status |
| **Automatic ODT Download** | Downloads ODT automatically if setup.exe is missing |
| **Internet Check** | Verifies connectivity before installation |

---

## Usage

### Portable Usage (Recommended)

1. Download `ActivaterWindows.exe`
2. Run it — that's it! **No installation required.**
3. Required files are downloaded and installed automatically
4. To remove, just delete `ActivaterWindows.exe` — **zero leftovers**

### Run from Source

```powershell
git clone https://github.com/tarikelertarnak/ActivaterWindows.git
cd ActivaterWindows
python builder/office_setup_wizard.py
```

### Build (Standalone .exe)

```powershell
powershell -ExecutionPolicy Bypass -File builder/build.ps1
```

The built executable is placed in the `dist/` folder.

---

## How to Use

### Tab 1: Office Installation
1. Choose **Installation Mode**: *Automatic* (LTSC 2024) or *Custom Configuration*
2. Select the installation language (e.g. `tr-tr`, `en-us`)
3. Press **Start Installation**
4. Office is downloaded, installed and (optionally) activated with Ohook

### Tab 2: Activation
1. Choose an **Activation Type** (combobox, default: *All — Windows + ESU + Office*):
   - `Windows 11 Pro — Permanent (HWID)`
   - `Windows 11 Pro for Workstations — Permanent (HWID)`
   - `Windows 11 Home / Windows 11 — Permanent (HWID)`
   - `Windows 10 Pro / Home / Windows 10 — Permanent (HWID)`
   - `Office LTSC/365 — Permanent (Ohook)`
   - `Windows 7-11 + Server — Permanent (TSforge)`
   - and more (ESU, Online KMS, uninstall)
2. Press **Activate**
3. Optionally verify with **Check Activation Status**

> **Edit Configuration Settings** is always available below the tabs — it is kept out of the installation flow on purpose.

### Cleanup

All temporary files are automatically removed when the application closes:
- `%TEMP%\ActivaterWindows\` folder
- `%TEMP%\oi_debug.log` debug log
- `%TEMP%\office_config.xml` configuration
- `%TEMP%\MAS_*.cmd` / `%TEMP%\CheckAct_*.cmd` activation scripts (deleted after use)

Deleting `ActivaterWindows.exe` leaves nothing behind.

---

## Technical Details

| Component | Version | Source |
|---|---|---|
| Office Deployment Tool | 16.0.20228.20124 | [aka.ms/odt](https://aka.ms/odt) |
| Microsoft Activation Scripts | v3.12 | [massgrave.dev](https://massgrave.dev) |
| Office LTSC Pro Plus 2024 | Volume License | PerpetualVL2024 channel |

### Supported Applications

Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher, Skype for Business, OneDrive, Visio, Project

### Supported Languages

38+ languages: Turkish, English, Deutsch, Français, Español, Italiano, Português, Nederlands, Polski, Русский, Українська, العربية, हिन्दी, Bahasa, ไทย, Tiếng Việt, 中文, 日本語, 한국어 and more.

### Activation Methods (MAS v3.12)

| Method | Scope | Duration |
|---|---|---|
| HWID | Windows 10/11 | Permanent (digital license) |
| TSforge | Windows 7-11 + Server | Permanent |
| ESU | Windows 10 (ESU) | Permanent |
| Ohook | Office LTSC/365 | Permanent |
| Online KMS | Windows / Office | 180 days (auto-renew task) |

---

## Project Structure

```
ActivaterWindows/
├── ActivaterWindows.exe                # Compiled portable application
├── builder/
│   ├── build.ps1                   # .exe build script
│   ├── office_setup_wizard.py      # Main application
│   ├── office_setup_wizard.spec    # PyInstaller spec (embeds MAS script)
│   ├── office_setup_wizard.ico     # Application icon
│   ├── version_info.txt            # Windows version info
│   └── icons/                      # Theme icons
│       ├── light/                  # Light theme icons
│       └── dark/                   # Dark theme icons
├── resources/
│   ├── setup.exe                   # Office Deployment Tool
│   ├── Configuration.xml           # Default configuration
│   ├── configuration-Office365-x64.xml
│   └── mas/
│       ├── MAS_AIO.cmd             # Embedded activation script (v3.12)
│       └── Check_Activation_Status.cmd
├── LICENSE                         # MIT License
└── README.md
```

---

## Disclaimer

This tool is built on top of official Microsoft tools (ODT). The activation feature uses the community-developed **Microsoft Activation Scripts (MAS)** project (GPL-3.0, [massgrave/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)). Using unlicensed software is subject to your country's laws and Microsoft's terms of use. This project is shared for educational and personal use only; usage responsibility rests entirely with the user.

---

## Contributing

Use the [Issues](https://github.com/tarikelertarnak/ActivaterWindows/issues) page for bug reports and feature requests.

**Made by:** [Tariq Eler](https://github.com/tarikelertarnak)

---

## License

MIT License — see the [LICENSE](LICENSE) file for details. The embedded MAS script comes from the GPL-3.0 licensed [Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts) project and is credited as such; the activation tab is a wrapper around that script.