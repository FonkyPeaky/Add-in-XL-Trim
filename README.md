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
- **Windows 11 rounded corners** via DWM API

---

## Prerequisites

| Requirement | Version |
|---|---|
| Windows | 10 or 11 |
| Microsoft Excel | 2016 / 2019 / 2021 / Microsoft 365 |
| .NET Framework | 4.7.2 |
| Visual Studio | 2022 (to build from source) |
| VSTO Runtime | [Download here](https://aka.ms/vsto40) |

---

## Installation

### Step 1 — Build the add-in

1. Clone the repository:
   ```
   git clone https://github.com/FonkyPeaky/Add-in-XL-Trim.git
   ```

2. Open `XLTrim.slnx` in **Visual Studio 2022**

3. Build the solution in **Debug** or **Release** configuration:
   - `Ctrl + Shift + B`

   Output goes to `XLTrim/bin/Debug/` or `XLTrim/bin/Release/`

---

### Step 2 — Register the add-in

Open `Install-XLTrim.reg` in a text editor and update the `Manifest` path to match your build output folder:

```reg
"Manifest"="C:\\Your\\Path\\To\\XLTrim\\bin\\Debug\\XLTrim.vsto|vstolocal"
```

Then double-click `Install-XLTrim.reg` and confirm the registry import.

> To uninstall, double-click `Uninstall-XLTrim.reg`.

---

### Step 3 — Open Excel

Restart Excel. The **XL Trim** button appears in the **Add-ins** tab on the ribbon.

---

## How to use

1. Open any `.xlsx` or `.xlsm` workbook in Excel
2. Go to the **Add-ins** tab → click **XL Trim**
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
| Named styles | 500 (only 20 used → 480 deletable) |
| Total named ranges | 3 000 |
| Invalid ranges | 2 500 (25 deleted sheets × 100 refs) |
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
├── ThisAddIn.cs                  # VSTO entry point
└── XLTrim.csproj
samples/
└── XL_Trim_Extreme_Test.xlsx    # Stress-test workbook
Install-XLTrim.reg               # Registry installer
Uninstall-XLTrim.reg             # Registry uninstaller
```

---

## License

MIT
