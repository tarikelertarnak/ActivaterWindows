# ActivaterWindows

> **Terminal-first Microsoft activator.** Windows, Office and ESU activation — one command, permanent methods, MAS power.

```powershell
irm https://tarikelertarnak.github.io/aw.ps1 | iex
```

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-3.0.2-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/size-%7E67KB-success?style=flat-square" />
  <img src="https://img.shields.io/badge/engine-auto--tracked%20%2B%20SHA256-red?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
  <a href="https://github.com/tarikelertarnak/ActivaterWindows/releases/latest"><img src="https://img.shields.io/github/v/release/tarikelertarnak/ActivaterWindows?style=flat-square&label=release" /></a>
</p>

**Made by TARIK ELER** — [tarikelertarnak.pages.dev](https://tarikelertarnak.pages.dev/) · [github.com/tarikelertarnak/ActivaterWindows](https://github.com/tarikelertarnak/ActivaterWindows)

---

## Why the Terminal Edition?

Version 1.0 was a single PyInstaller `.exe` (19 MB). Version 2.x rebuilt the whole foundation as a plain-text PowerShell script:

| | v1.0 (GUI exe) | v3.0.1 (terminal) |
|---|---|---|
| Size | 19 MB PE binary | **~67 KB** of plain text |
| Format | Compiled, packed, unreadable | **Every line readable**, no packer, no compiler needed |
| Startup | PyInstaller unpacking | Instant |
| MAS features | Embedded, frozen at v3.12 | Auto-tracked upstream + two fallback tiers — **all features** (Troubleshoot, edition change, `$OEM$`, and more) |
| Antivirus heuristics | High (packed exe + license API calls are a classic heuristic trap) | Much lower (nothing packed, signed remote sources, SHA256 verification) |

Antivirus note: the script uses a plain-text approach with integrity verification across three signed sources, and follows the activation engine's own official download method. Because there is no packed binary, reputation-based scanners flag it far less than an exe. If a third-party antivirus still blocks downloading the script, the script detects this and tells you.

---

## Usage

### One-liner (recommended)

Open PowerShell, paste:

```powershell
irm https://tarikelertarnak.github.io/aw.ps1 | iex
```

TLS 1.2 is enforced explicitly at the very top of the script, so machines whose defaults would otherwise fail the connection work without any extra setup. The menu appears; UAC prompts when administrator rights are needed.

### Local file

1. Download `ActivaterWindows.cmd` and `ActivaterWindows.ps1` (keep them in the same folder)
2. Double-click `ActivaterWindows.cmd`

Or launch the script directly:

```powershell
powershell -ExecutionPolicy Bypass -File ActivaterWindows.ps1
```

### Runs on any clean Windows 10/11 machine

Nothing has to be pre-installed by the user. A single terminal command is enough.

| | |
|---|---|
| Required | Windows 10/11, internet access, administrator rights (UAC) |
| Pre-installed by you | Nothing |
| Written to disk | A temporary `.cmd` plus its log; both are deleted when the operation finishes. No permanent files |
| Older Windows | Windows 7/8 are unsupported (TLS 1.2 is required) |

Because the one-liner runs through `irm | iex`, no `ExecutionPolicy` change is needed either.

---

## Menu

The **language selector sits at the very top** of the interface: **[A] Auto / [T] Türkçe / [E] English**.

| # | Feature | Duration |
|---|---|---|
| 1 | **Activate Windows** — opens a submenu (see below) | Permanent |
| 2 | **Activate Office** — permanent Ohook-based activation | Permanent |
| 3 | **Windows + ESU + Office** — one pass; Office activation runs only if Office is already installed | Permanent |
| 4 | **Windows 10 ESU** | Permanent |
| 5 | **Check activation status** | — |
| 6 | **Install Office LTSC Professional Plus 2024** — language first, then mode/components | — |
| 7 | **Activate everything** — activates what is already installed; **no Office install** | Permanent |
| 8 | **Do everything** — installs Office if missing, activates everything, then offers the full engine menu | Permanent |
| 9 | **About** — version, author, license attribution, links | — |
| 0 | **Exit** | — |

**Esc** goes back from any submenu. Picking the wrong entry never forces you to restart the app — just step back and choose again.

> The menu only says **what** it does, never **which method** it uses. Method names never appear in the interface.
> Method argument ordering does not matter: arguments are sent in a single pass and all of them are applied regardless of the order the engine processes them.

### [1] Activate Windows submenu

- Show the current edition
- List the available editions (read from DISM/WMI) and change one, then activate
- Windows 10 ESU

### Office activation runs hidden

Activation runs in the background with **no window opened**. The engine's own output goes to a temporary log; only the result lines are filtered and shown:

```
  Windows 11 Pro is LICENSED.
  Office is permanently activated.
```

---

## Installing Office LTSC Professional Plus 2024

The Office install flow always downloads the **Office Deployment Tool first**, and only then asks you questions:

1. **ODT download, in order of attempts** — local copy → Microsoft's download page → pinned Microsoft URL → GitHub release mirror (`odt-mirror`). Each candidate is verified by **file size + PE header** before use.
2. **Language** — 83 Office language codes are listed alphabetically with numbers:
   - Type a **code** such as `tr-tr` → accepted directly
   - Type a **number** (e.g. `76`) → confirmation: "76) tr-tr — confirm? Enter = yes, 0 = back"
   - **Empty Enter** → the auto-detected language is used
3. **Mode / component selection** — see below.

### Component picker

| Mode | What it installs |
|---|---|
| **1 — Recommended (light)** | Word, Excel, PowerPoint only (the other 6 components are not installed) |
| **2 — Full suite** | All components |
| **3 — Custom selection** | Tick the 9 components one by one |

The Custom selection picker opens with **nothing pre-selected** — you decide what you want. It shows a live counter:

```
  ==========================================================
   CUSTOM SELECTION                             Selected: 0 / 9
  ==========================================================
    1) [ ] Word               recommended
    2) [ ] Excel              recommended
    3) [ ] PowerPoint         recommended
    4) [ ] Outlook
    5) [ ] Access
    6) [ ] Publisher
    7) [ ] OneNote
    8) [ ] OneDrive
    9) [ ] Skype for Business

  Toggle: comma-separated numbers (example: 1,3,5)   |   a = all   n = none   r = recommended (Word, Excel, PowerPoint)   Enter = confirm
```

- Write numbers separated by commas to toggle them: `1,3,5`
- `a` = all · `n` = none · `r` = recommended (Word, Excel, PowerPoint) · **Enter** = confirm
- At least one component is required
- After confirmation the **names of the components to be installed** are printed as well (e.g. `Installing: Word, Excel, PowerPoint`).

### Progress bars

Jobs with a known size (downloads) show a **real percentage**; long steps with unknown duration show a **sweeping bar plus elapsed seconds**:

```
  [############------------------]  40%    3s     <- download (real percentage)
  [-----------#-----------------]    7s            <- ODT extraction / Office setup / activation
```

The sweeping bar runs in three places: **ODT extraction**, **Office setup** (`setup.exe /configure`) and **hidden activation**. All three run windowless, so this bar is the only feedback — previously a single line of text appeared and the screen just froze.

Downloads always show a real percentage bar with a seconds counter:

```
  [############------------------]  40%    3s
```

### Proof / screenshots

Screenshots below were taken from a real run and show the activation result the script prints:

- `assets/proof/office-activated.png` — Office activation result (`Office is permanently activated`)
- `assets/proof/activation-status.png` — status check showing the current edition and build

Drop your own screenshots into `assets/proof/` and reference them here; they are not part of the script download. Because hidden activation writes its output to a temporary log that is deleted afterwards, a screenshot is the only way to keep a record of the result on your own machine.

---

## How it works

### Engine resolution (3 tiers, all SHA256-verified)

| Tier | Source | Verification | Why |
|---|---|---|---|
| 1 | Upstream's own bootstrap (`get.activated.win`) → the current commit + hash | SHA256 read from the bootstrap | The engine stays current even when the app itself is not updated |
| 2 | A pin embedded in the script body (commit + hash) | Embedded SHA256 | Keeps working if upstream changes its format or argument order |
| 3 | Local cache (`%LOCALAPPDATA%\ActivaterWindows`) | SHA256 recorded at cache time | Keeps working even if the network or antivirus blocks everything |

- **SHA256 verification:** in tiers 1 and 2, content whose hash does not match is **never executed** — a corrupted or tampered download is rejected and the script falls to the next tier.
- **3 mirrors:** raw.githubusercontent → Azure DevOps → git.activated.win, tried in random order.
- **Temporary file:** the executed cmd is written to a temp folder (administrator: `%SystemRoot%\Temp`, otherwise `%TEMP%`) and deleted when the job finishes.

> **Honest limit (tier 3):** the local cache is verified against the `.sha256` file next to it. That catches corruption and later modification, but it does not stop an attacker running as the same user from changing both files together. An actor who already has administrator rights is outside this scope anyway. This tier is a resilience fallback, not a trust boundary.

### ODT resolution (4 attempts, in order)

1. A local copy
2. Microsoft's download page
3. A pinned Microsoft URL
4. The GitHub release mirror `odt-mirror`

Every candidate is verified by **file size + PE header** before it is used.

> **License note:** the bundled activation engine is the GPL-3.0 licensed **Microsoft Activation Scripts**. Attribution is kept in the README and disclaimer for license compliance; it is never shown in the interface.

---

## Security architecture

- **Nothing runs without a verified hash** in tiers 1 and 2, and tier 3 is re-verified on every use.
- **Hidden activation is a privacy-only filter, not a trust boundary.** Method and brand names, source addresses, license texts and raw cmd echoes are suppressed; result and warning lines are kept.
- **UAC stays visible.** It is a security boundary and cannot be hidden. Nothing runs until you approve it.
- **Timeout on hidden runs.** Some engine flows wait for a keypress at the end (`choice /c 0 /n`). In a hidden window that key can never arrive, so the process would hang forever. Hidden runs therefore get a **900-second** cap: when it expires the process is closed, whatever results reached the log are still shown, and you are warned that the operation may not have completed. If the filter leaves no lines at all, a "no result lines" message is printed — so **an empty screen is never confused with a failure**.
- **Built-in checks:** `-SelfTest` validates the filter, language parity and brand leakage without needing network or admin rights; `-Probe` downloads the engine and verifies its hash. Neither of them runs anything.

---

## What's new in v2.5.0

- **Version bumped to 2.5.0** — polish and finalisation pass (bug fixes, UI tidy-up, proof section)
- Everything from v2.3.0 stays: Esc back, Windows editions submenu, About screen, made-by credit, shortened one-liner, ODT mirror fallback

The changes from v2.3.0:


## Limitations

- **Windows Defender and other security tools may flag activation tools.** This is expected for software in this category.
- **Installation is capped at 60 minutes.** A longer install is treated as a failure rather than being allowed to run indefinitely.
- **Windows 7/8 are unsupported** — TLS 1.2 is required.
- **Offline use is partial.** After a first successful run the engine is cached, verified, in `%LOCALAPPDATA%\ActivaterWindows`; with no network that copy is used. With no cache *and* no network, activation cannot proceed, because integrity verification is required.
- **Third-party antivirus may block the script download.** Add the folder to your antivirus exclusions, or use the `irm … | iex` one-liner (nothing is written to disk; the download is verified and handed to cmd).

---

## Project structure

```
ActivaterWindows/
├── ActivaterWindows.ps1      # Terminal core (single file, ~67 KB)
├── ActivaterWindows.cmd      # Double-click launcher
├── builder/                  # (legacy) v1.0 GUI exe source code
├── resources/                # (legacy) embedded ODT + MAS v3.12 copies
└── mas-src/                  # (reference) MAS source snapshot
```

## Technical details

| Component | Source |
|---|---|
| Activation engine | [Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts) (GPL-3.0); current commit tracked from upstream, with a pinned commit as fallback |
| Office Deployment Tool | [aka.ms/odt](https://aka.ms/odt) (downloaded on demand at install time) |
| Office LTSC Pro Plus 2024 | PerpetualVL2024 channel, x64 |

---

## Old version (GUI exe — archived)

The v1.0 Tkinter GUI (`builder/`, `resources/`) is kept as legacy, but it is **no longer developed** and, being a packed `.exe`, it can trigger antivirus heuristics. It is archived only:
[v1.0.0 (archive)](https://github.com/tarikelertarnak/ActivaterWindows/releases/tag/v1.0.0).

**The version you should use:** [v2.3.0 — Terminal Edition](https://github.com/tarikelertarnak/ActivaterWindows/releases/tag/v2.3.0)
(the one-liner above). You do not need to download `ActivaterWindows.exe`.

---

## FAQ

**My antivirus warned me — what should I do?** Windows Defender usually leaves it alone. If a third-party antivirus blocks the script download: add the folder to your antivirus exclusions, or use the `irm … | iex` one-liner (nothing is written to disk; the download is verified and handed to cmd).

**Does it work offline?** Partly. After a first successful run the engine is cached, verified, in `%LOCALAPPDATA%\ActivaterWindows`; with no network that copy is used. With no cache *and* no network, activation cannot proceed, because integrity verification is required.

---

## Disclaimer

This project is a wrapper around the community-developed **Microsoft Activation Scripts (MAS)** (GPL-3.0, [massgrave/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)). Using unlicensed software is subject to your country's laws and Microsoft's terms of use. Educational and personal use only; usage responsibility rests entirely with the user.

**Made by:** [TARIK ELER](https://tarikelertarnak.pages.dev/) · [github.com/tarikelertarnak](https://github.com/tarikelertarnak)

## License

MIT — see [LICENSE](LICENSE). MAS is GPL-3.0 and credited above.
