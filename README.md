# Contact Migrator

This repository contains one simple bootstrap script that can create and run the full project.

## What this does

The project provides an interactive desktop GUI that:
- lets you browse for files in Finder on macOS or File Explorer on Windows
- accepts:
  - Apple / iPhone Contacts `.vcf` or `.csv`
  - Apple / iPhone Calendar birthdays `.ics`
  - Google Calendar export `.ics`
  - Outlook Contacts `.csv`
- creates a workspace with:
  - `input/`
  - `results/`
- copies selected inputs into `input/`
- creates:
  - `combined_contacts.csv`
  - `combined_contacts.xlsx` if `openpyxl` is installed
  - `iphone_contacts_import.vcf`
  - `outlook_contacts_import.csv`
  - `google_birthdays_import.ics`
  - `summary.txt`

## The only script you need first

Use:

- `bootstrap_contact_migrator.py`

That script can:
- explain what the project does
- write all required project files into this repository
- optionally launch the GUI

## Quick start on macOS

```bash
python3 bootstrap_contact_migrator.py --write
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 bootstrap_contact_migrator.py --run-mac
```

## Quick start on Windows

```powershell
py bootstrap_contact_migrator.py --write
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
py bootstrap_contact_migrator.py --run-windows
```

## What `--write` creates

```text
contact-migrator/
  README.md
  requirements.txt
  .gitignore
  index.md
  bootstrap_contact_migrator.py
  contact_migrator/
    __init__.py
    core.py
    gui.py
  scripts/
    run_mac.command
    run_windows.bat
```

## Commands

Show help and explanation:

```bash
python3 bootstrap_contact_migrator.py --help
```

Write the full project files:

```bash
python3 bootstrap_contact_migrator.py --write
```

Run the GUI for macOS:

```bash
python3 bootstrap_contact_migrator.py --run-mac
```

Run the GUI for Windows:

```bash
py bootstrap_contact_migrator.py --run-windows
```

## GitHub setup later

After the files are written and tested locally:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/contact-migrator.git
git push -u origin main
```

## Notes

- This version does not include GitHub workflow files.
- Users choose files manually with a file picker, so filenames do not need to follow any naming convention.
- The script keeps going even if some sources are skipped.
