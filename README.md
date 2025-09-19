# Folder Exporter with Progress UI — README

## Overview

This Python program scans a selected folder (Dropbox, OneDrive, or local) and exports all subfolders into a structured spreadsheet (Excel `.xlsx` or `.csv`). Each folder is listed on its own row, with parent folders repeated across columns (`Level1`, `Level2`, …).

It works with online‑only synced folders (Dropbox/OneDrive Files On‑Demand), since it only reads folder names and never downloads file contents.

## Key Features

* GUI folder picker
* Save location prompt before scanning
* Progress window with folder count and current path
* Cancel button (partial results still saved)
* Windows long‑path (`\\?\`) support
* Excel `.xlsx` output (auto‑installs `openpyxl`) or CSV fallback
* Skipped/inaccessible folders logged to `*.scan_log.txt`
* Compatible with online‑only synced folders

## Installation

### Requirements

* Python 3.8+
* `tkinter` (bundled with Python on Windows/macOS)
* `openpyxl` (auto‑installed if missing, for Excel support)

### Download the script file

```
export_folders_progress_ui_winfix.py
```

### (Optional) Manually install `openpyxl`

```bash
pip install openpyxl
```

## Usage

Run the script:

```bash
python export_folders_progress_ui_winfix.py
```

Then:

1. Choose the folder to scan (Dropbox, OneDrive, or local).
2. Select where to save the output file (`.xlsx` or `.csv`).
3. On Windows, confirm whether to enable long‑path prefix support (`\\?\`).
4. Monitor the progress window:

   * Shows folders found
   * Displays current scanning path
   * Cancel button stops early (partial results saved)

## Output

### Spreadsheet

* Each row = one folder
* Columns = folder depth (`Level1`, `Level2`, …)

### Log file

* Named `yourfile.scan_log.txt`
* Lists skipped folders with reasons (for example, moved during scan, permission errors)

## Example

```
Level1   | Level2     | Level3
---------|------------|--------
Dropbox  | Projects   |
Dropbox  | Projects   | Alpha
Dropbox  | Projects   | Alpha | Docs
Dropbox  | Projects   | Beta
```

## Notes

* Works with Dropbox and OneDrive online‑only synced folders.
* Script never downloads files; it only traverses folder entries exposed by the sync client.
* Online‑only folders that the sync client does not yet expose may still be skipped.
* For very deep folder trees, enable long‑path support on Windows.
