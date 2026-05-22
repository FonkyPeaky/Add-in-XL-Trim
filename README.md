# XL Trim — Excel Add-in

XL Trim is a lightweight Excel VSTO add-in that scans and cleans bloated workbooks in one click.

It removes unused named styles, deletes invalid named ranges, and restores hidden ranges — directly from an Excel ribbon button, with a live animated progress ring and health score.

---

## What it fixes

| Problem | What XL Trim does |
|---|---|
| Thousands of unused named styles | Scans every cell to find styles in use, deletes the rest |
| Invalid named ranges (`#REF!`, missing sheets) | Detects and deletes broken references |
| Hidden named ranges (filter caches, add-in leftovers) | Makes them all visible again |

---

## Features

- **Scan dialog** — shows styles count, named ranges, invalid count, hidden count, and a health score (0–100%)
- **Animated progress ring** — smooth 60 fps easing during cleanup
- **Backup first** — SaveFileDialog prompts for a backup location (defaults to Downloads) before any change
- **Fast detection** — invalid ranges pointing to deleted sheets are resolved via a HashSet lookup, not COM exceptions — handles 50 000+ ranges without freezing
- **No network activity** — everything runs locally, the file never leaves the user's machine
- **Windows 11 rounded corners** via DWM API

---

## Prerequisites

| Requirement | Version |
|---|---|
| Windows | 10 or 11 |
| Microsoft Excel | 2016 / 2019 / 2021 / Microsoft 365 |
| .NET Framework | 4.7.2 |
| VSTO Runtime | [Download here](https://aka.ms/vsto40) |
| Visual Studio | 2022 (to build from source only) |

---

## Installation

### Step 1 — Build the add-in

1. Clone the repository:
   ```
   git clone https://github.com/FonkyPeaky/Add-in-XL-Trim.git
   ```

2. Open `XLTrim.slnx` in **Visual Studio 2022**

3. Build in **Release** configuration (`Ctrl + Shift + B`)

   Output: `XLTrim/bin/Release/`

---

### Step 2 — Install

Run `Install.ps1` as administrator from the `bin/Release/` folder:

```powershell
powershell -ExecutionPolicy Bypass -File Install.ps1
```

This copies the add-in to `C:\Program Files\XLTrim\` and registers it in `HKLM` for all users on the machine.

> To uninstall: `powershell -ExecutionPolicy Bypass -File Uninstall.ps1`

---

### Step 3 — Open Excel

Restart Excel. The **XL Trim** tab appears in the ribbon with the **Clean File** button.

---

## Software Center / SCCM deployment

To deploy via SCCM or Software Center:

1. Build the solution in **Release** configuration
2. Package the contents of `bin\Release\` together with `Install.ps1` and `Uninstall.ps1`
3. Set the deployment commands:

| Action | Command |
|---|---|
| Install | `powershell -ExecutionPolicy Bypass -File Install.ps1` |
| Uninstall | `powershell -ExecutionPolicy Bypass -File Uninstall.ps1` |
| Detection | Key exists: `HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\XLTrim` |

The script registers the add-in machine-wide (`HKLM`) and covers both 64-bit and 32-bit Excel registry hives automatically.

**Prerequisites to deploy separately:**
- .NET Framework 4.7.2 (included in Windows 10 1803+ and Windows 11)
- [VSTO Runtime 4.0](https://aka.ms/vsto40)

---

## How to use

1. Open any `.xlsx` or `.xlsm` workbook in Excel
2. Go to the **XL Trim** tab → click **Clean File**
3. Choose whether to create a backup (recommended)
4. The scan dialog opens and analyzes the workbook
5. Review the results — styles count, invalid ranges, hidden ranges, health score
6. Click **Clean File** to remove all issues
7. The file is saved automatically when done

---

## Health score

| Score | Label | Meaning |
|---|---|---|
| 80–100% | Good | File is clean |
| 50–79% | Average | Some issues present |
| 0–49% | Critical | Severely bloated |

Deductions:
- **−20 pts** if Styles Max ID > 4 000 (−40 if > 64 000)
- **−15 pts** if any invalid named ranges exist (−30 if > 1 000)
- **−10 pts** if more than 100 hidden named ranges

---

## Sample test file

A stress-test workbook is included in the [`samples/`](samples/) folder:

**`XL_Trim_Extreme_Test.xlsx`**

| Metric | Value |
|---|---|
| Named styles | 500 (only 20 used — 480 deletable) |
| Total named ranges | 3 000 |
| Invalid ranges | 2 500 (25 deleted sheets x 100 refs) |
| Hidden ranges | 400 |
| Health score | Critical |

Open it in Excel, launch XL Trim, and observe the full scan + clean cycle.

---

## Project structure

```
XLTrim/
├── Core/
│   ├── WorkbookScanner.cs       # Scans styles, named ranges
│   ├── StylesCleaner.cs         # Removes unused named styles
│   ├── NamedRangesCleaner.cs    # Deletes invalid, unhides hidden ranges
│   ├── BackupManager.cs         # File copy helper
│   └── ScanResult.cs            # Result data model
├── Ribbon/
│   ├── XLTrimRibbon.cs          # Ribbon button handler, backup flow
│   └── XLTrimRibbon.xml         # Ribbon XML definition
├── UI/
│   ├── GlassWindow.cs           # Base window with Win11 rounded corners
│   ├── ScanDialog.xaml/.cs      # Scan + clean dialog
│   └── BackupDialog.xaml/.cs    # Backup prompt dialog
├── broom.png                     # Ribbon button icon (embedded resource)
├── ThisAddIn.cs                  # VSTO entry point
└── XLTrim.csproj
samples/
└── XL_Trim_Extreme_Test.xlsx    # Stress-test workbook
Install.ps1                       # Installer (SCCM / Software Center ready)
Uninstall.ps1                     # Uninstaller
```

---

## License

MIT
