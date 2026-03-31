#!/usr/bin/env python3
import argparse
import subprocess
import sys
from pathlib import Path

README_MD = """# Contact Migrator

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
.venv\\Scripts\\Activate.ps1
pip install -r requirements.txt
py bootstrap_contact_migrator.py --run-windows
```
"""

REQUIREMENTS_TXT = "openpyxl\n"

GITIGNORE = """.venv/
__pycache__/
*.pyc
.DS_Store
Thumbs.db
*.xlsx
*.log
"""

INDEX_MD = """# Contact Migrator

Use the repository README for setup and usage.
"""

INIT_PY = "__all__ = []\n"

CORE_PY = r"""
import csv
import re
import shutil
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except Exception:
    openpyxl = None


def clean_text(value):
    if value is None:
        return ""
    value = str(value).replace("\\r", " ").replace("\\n", " ").strip()
    value = re.sub(r"\\s+", " ", value)
    return value


def ensure_workspace(project_dir):
    project_dir = Path(project_dir)
    input_dir = project_dir / "input"
    results_dir = project_dir / "results"
    input_dir.mkdir(parents=True, exist_ok=True)
    results_dir.mkdir(parents=True, exist_ok=True)
    return project_dir, input_dir, results_dir


def copy_if_present(src, dst_dir):
    if not src:
        return None
    src_path = Path(src)
    if not src_path.exists():
        return None
    dst = Path(dst_dir) / src_path.name
    if src_path.resolve() != dst.resolve():
        shutil.copy2(src_path, dst)
    return dst


def normalize_header(h):
    return re.sub(r"[^a-z0-9]+", "_", h.strip().lower()).strip("_")


def find_column(fieldnames, candidates):
    normalized = {normalize_header(f): f for f in fieldnames}
    for c in candidates:
        if c in normalized:
            return normalized[c]
    return None


def split_multi(value):
    value = clean_text(value)
    if not value:
        return []
    parts = re.split(r"\\s*\\|\\s*|\\s*;\\s*(?=\\S)", value)
    return [clean_text(p) for p in parts if clean_text(p)]


def strip_label_suffix(value):
    return re.sub(r"\\s*\\[[^\\]]+\\]\\s*$", "", clean_text(value)).strip()


def unfold_vcf_lines(file_obj):
    buf = None
    for raw in file_obj:
        line = raw.rstrip("\\r\\n")
        if buf is None:
            buf = line
            continue
        if line.startswith((" ", "\\t")):
            buf += line[1:]
        else:
            yield buf
            buf = line
    if buf is not None:
        yield buf


def split_prop(line):
    in_quotes = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == ":" and not in_quotes:
            return line[:i], line[i + 1:]
    return line, ""


def parse_params(header):
    parts = header.split(";")
    name = parts[0].upper()
    params = {}
    for part in parts[1:]:
        if "=" in part:
            k, v = part.split("=", 1)
            params[k.upper()] = [x.strip().strip('"') for x in v.split(",")]
        else:
            params.setdefault("TYPE", []).append(part.strip().strip('"'))
    return name, params


def unescape_value(value):
    return (
        value.replace("\\\\n", "\\n")
        .replace("\\\\N", "\\n")
        .replace("\\\\,", ",")
        .replace("\\\\;", ";")
        .replace("\\\\\\\\", "\\\\")
    )


def parse_vcf(path):
    contacts = []
    current = None

    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        for line in unfold_vcf_lines(f):
            line = line.strip()
            if not line:
                continue
            upper = line.upper()
            if upper == "BEGIN:VCARD":
                current = {
                    "full_name": "",
                    "first_name": "",
                    "middle_name": "",
                    "last_name": "",
                    "prefix": "",
                    "suffix": "",
                    "organization": "",
                    "title": "",
                    "emails": [],
                    "phones": [],
                    "addresses": [],
                    "birthday": "",
                    "note": "",
                    "source_type": "vcf",
                }
                continue
            if upper == "END:VCARD":
                if current is not None:
                    if not current["full_name"]:
                        parts = [
                            current["prefix"],
                            current["first_name"],
                            current["middle_name"],
                            current["last_name"],
                            current["suffix"],
                        ]
                        current["full_name"] = " ".join([p for p in parts if p]).strip()
                    contacts.append(current)
                current = None
                continue
            if current is None:
                continue

            header, value = split_prop(line)
            prop, _params = parse_params(header)
            value = unescape_value(value)

            if prop == "FN":
                current["full_name"] = clean_text(value)
            elif prop == "N":
                parts = value.split(";")
                while len(parts) < 5:
                    parts.append("")
                current["last_name"] = clean_text(parts[0])
                current["first_name"] = clean_text(parts[1])
                current["middle_name"] = clean_text(parts[2])
                current["prefix"] = clean_text(parts[3])
                current["suffix"] = clean_text(parts[4])
            elif prop == "EMAIL":
                current["emails"].append(clean_text(value))
            elif prop == "TEL":
                current["phones"].append(clean_text(value))
            elif prop == "ADR":
                adr_parts = value.split(";")
                street = clean_text(adr_parts[2]) if len(adr_parts) > 2 else ""
                city = clean_text(adr_parts[3]) if len(adr_parts) > 3 else ""
                region = clean_text(adr_parts[4]) if len(adr_parts) > 4 else ""
                postal = clean_text(adr_parts[5]) if len(adr_parts) > 5 else ""
                country = clean_text(adr_parts[6]) if len(adr_parts) > 6 else ""
                addr = ", ".join([x for x in [street, city, region, postal, country] if x])
                if addr:
                    current["addresses"].append(addr)
            elif prop == "ORG":
                current["organization"] = clean_text(value.replace(";", " - "))
            elif prop == "TITLE":
                current["title"] = clean_text(value)
            elif prop == "BDAY":
                current["birthday"] = clean_text(value)
            elif prop == "NOTE":
                current["note"] = clean_text(value)

    return contacts


def parse_contacts_csv(path):
    contacts = []
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            return contacts

        fieldnames = reader.fieldnames
        colmap = {
            "full_name": find_column(fieldnames, ["full_name", "name", "display_name", "fn", "contact_name"]),
            "first_name": find_column(fieldnames, ["first_name", "firstname", "given_name"]),
            "middle_name": find_column(fieldnames, ["middle_name", "middlename"]),
            "last_name": find_column(fieldnames, ["last_name", "lastname", "family_name", "surname"]),
            "prefix": find_column(fieldnames, ["prefix"]),
            "suffix": find_column(fieldnames, ["suffix"]),
            "organization": find_column(fieldnames, ["organization", "org", "company"]),
            "title": find_column(fieldnames, ["title", "job_title"]),
            "emails": find_column(fieldnames, ["emails", "email", "mail"]),
            "phones": find_column(fieldnames, ["phones", "phone", "telephone", "mobile", "tel"]),
            "addresses": find_column(fieldnames, ["addresses", "address"]),
            "birthday": find_column(fieldnames, ["birthday", "bday", "birth_date", "birthday_date", "dob", "date_of_birth"]),
            "note": find_column(fieldnames, ["note", "notes"]),
        }

        for row in reader:
            contact = {
                "full_name": clean_text(row.get(colmap["full_name"], "")) if colmap["full_name"] else "",
                "first_name": clean_text(row.get(colmap["first_name"], "")) if colmap["first_name"] else "",
                "middle_name": clean_text(row.get(colmap["middle_name"], "")) if colmap["middle_name"] else "",
                "last_name": clean_text(row.get(colmap["last_name"], "")) if colmap["last_name"] else "",
                "prefix": clean_text(row.get(colmap["prefix"], "")) if colmap["prefix"] else "",
                "suffix": clean_text(row.get(colmap["suffix"], "")) if colmap["suffix"] else "",
                "organization": clean_text(row.get(colmap["organization"], "")) if colmap["organization"] else "",
                "title": clean_text(row.get(colmap["title"], "")) if colmap["title"] else "",
                "emails": split_multi(row.get(colmap["emails"], "")) if colmap["emails"] else [],
                "phones": split_multi(row.get(colmap["phones"], "")) if colmap["phones"] else [],
                "addresses": split_multi(row.get(colmap["addresses"], "")) if colmap["addresses"] else [],
                "birthday": clean_text(row.get(colmap["birthday"], "")) if colmap["birthday"] else "",
                "note": clean_text(row.get(colmap["note"], "")) if colmap["note"] else "",
                "source_type": "csv",
            }
            if not contact["full_name"]:
                parts = [contact["prefix"], contact["first_name"], contact["middle_name"], contact["last_name"], contact["suffix"]]
                contact["full_name"] = " ".join([p for p in parts if p]).strip()
            contacts.append(contact)

    return contacts


def unfold_ics_lines(lines):
    buffer = ""
    for raw in lines:
        line = raw.rstrip("\\r\\n")
        if line.startswith((" ", "\\t")):
            buffer += line[1:]
        else:
            if buffer:
                yield buffer
            buffer = line
    if buffer:
        yield buffer


def parse_ics_property(line):
    if ":" not in line:
        return line, "", {}
    left, value = line.split(":", 1)
    parts = left.split(";")
    name = parts[0].upper()
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            params[k.upper()] = v
        else:
            params[p.upper()] = True
    return name, value, params


def normalize_date_string(raw):
    raw = clean_text(raw)
    if not raw:
        return ""
    raw = raw.replace("/", "-").replace(".", "-")
    m = re.match(r"^(\\d{4})-(\\d{2})-(\\d{2})", raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.match(r"^(\\d{8})$", raw)
    if m:
        s = m.group(1)
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return raw


def extract_name_from_summary(summary):
    s = clean_text(summary)
    patterns = [
        r"^(.*?)'?s Birthday$",
        r"^(.*?)’s Birthday$",
        r"^Birthday: (.*)$",
        r"^(.*) Birthday$",
        r"^BD-(.*?)(?: \\d{4}-\\d{2}-\\d{2})?$",
    ]
    for pat in patterns:
        m = re.match(pat, s, flags=re.IGNORECASE)
        if m:
            return clean_text(m.group(1))
    return s


def parse_birthdays_ics(path):
    contacts = []
    current = None
    in_event = False
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in unfold_ics_lines(f):
            line = line.strip()
            if line == "BEGIN:VEVENT":
                current = {}
                in_event = True
                continue
            if line == "END:VEVENT":
                if in_event and current:
                    summary = clean_text(current.get("SUMMARY", ""))
                    categories = clean_text(current.get("CATEGORIES", "")).lower()
                    description = clean_text(current.get("DESCRIPTION", "")).lower()
                    if "birthday" in summary.lower() or "birthday" in categories or "birthday" in description or summary.startswith("BD-"):
                        name = extract_name_from_summary(summary)
                        dt = normalize_date_string(current.get("DTSTART", ""))
                        contacts.append({
                            "full_name": name,
                            "first_name": "",
                            "middle_name": "",
                            "last_name": "",
                            "prefix": "",
                            "suffix": "",
                            "organization": "",
                            "title": "",
                            "emails": [],
                            "phones": [],
                            "addresses": [],
                            "birthday": dt,
                            "note": "",
                            "source_type": "ics_birthday",
                        })
                current = None
                in_event = False
                continue

            if not in_event or current is None:
                continue

            name, value, _params = parse_ics_property(line)
            if name in {"SUMMARY", "DESCRIPTION", "CATEGORIES", "UID", "DTSTART"}:
                current[name] = value
    return contacts


def dedupe_contacts(contacts):
    deduped = {}
    for c in contacts:
        name = clean_text(c.get("full_name", "")).lower()
        bday = clean_text(c.get("birthday", "")).lower()
        email = clean_text(c.get("emails", [""])[0] if c.get("emails") else "").lower()
        phone = clean_text(c.get("phones", [""])[0] if c.get("phones") else "").lower()

        if name and bday:
            key = ("name_bday", name, bday)
        elif email:
            key = ("email", email)
        elif phone:
            key = ("phone", phone)
        elif name:
            key = ("name", name)
        else:
            key = ("row", id(c))

        if key not in deduped:
            deduped[key] = c
        else:
            existing = deduped[key]
            for field in ["full_name", "first_name", "middle_name", "last_name", "prefix", "suffix", "organization", "title", "birthday", "note"]:
                if not existing.get(field) and c.get(field):
                    existing[field] = c[field]
            for list_field in ["emails", "phones", "addresses"]:
                combined = existing.get(list_field, []) + c.get(list_field, [])
                seen = set()
                out = []
                for item in combined:
                    v = clean_text(item)
                    if v and v.lower() not in seen:
                        seen.add(v.lower())
                        out.append(v)
                existing[list_field] = out
    return list(deduped.values())


def write_combined_csv(contacts, out_path):
    fields = [
        "full_name", "first_name", "middle_name", "last_name", "prefix", "suffix",
        "organization", "title", "emails", "phones", "addresses", "birthday", "note", "source_type"
    ]
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for c in contacts:
            row = dict(c)
            row["emails"] = " | ".join(c.get("emails", []))
            row["phones"] = " | ".join(c.get("phones", []))
            row["addresses"] = " | ".join(c.get("addresses", []))
            writer.writerow(row)


def write_xlsx_if_possible(csv_path, xlsx_path):
    if openpyxl is None:
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"
    with open(csv_path, "r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)
    wb.save(xlsx_path)
    return True


def vcf_escape(value):
    value = clean_text(value)
    return value.replace("\\\\", "\\\\\\\\").replace(";", r"\\;").replace(",", r"\\,").replace("\\n", r"\\n")


def write_iphone_vcf(contacts, out_path):
    with open(out_path, "w", encoding="utf-8", newline="\\r\\n") as f:
        for c in contacts:
            f.write("BEGIN:VCARD\\r\\n")
            f.write("VERSION:3.0\\r\\n")
            f.write(f"N:{vcf_escape(c.get('last_name',''))};{vcf_escape(c.get('first_name',''))};{vcf_escape(c.get('middle_name',''))};{vcf_escape(c.get('prefix',''))};{vcf_escape(c.get('suffix',''))}\\r\\n")
            f.write(f"FN:{vcf_escape(c.get('full_name',''))}\\r\\n")
            if c.get("organization"):
                f.write(f"ORG:{vcf_escape(c['organization'])}\\r\\n")
            if c.get("title"):
                f.write(f"TITLE:{vcf_escape(c['title'])}\\r\\n")
            if c.get("birthday"):
                bday = normalize_date_string(c["birthday"])
                if re.match(r"^\\d{4}-\\d{2}-\\d{2}$", bday):
                    f.write(f"BDAY:{bday}\\r\\n")
            for email in c.get("emails", []):
                f.write(f"EMAIL;TYPE=INTERNET:{vcf_escape(strip_label_suffix(email))}\\r\\n")
            for phone in c.get("phones", []):
                f.write(f"TEL:{vcf_escape(strip_label_suffix(phone))}\\r\\n")
            for addr in c.get("addresses", []):
                f.write(f"ADR;TYPE=HOME:;;{vcf_escape(strip_label_suffix(addr))};;;;\\r\\n")
            if c.get("note"):
                f.write(f"NOTE:{vcf_escape(c['note'])}\\r\\n")
            f.write("END:VCARD\\r\\n")


def write_outlook_csv(contacts, out_path):
    fields = [
        "First Name", "Middle Name", "Last Name", "Title", "Company",
        "E-mail Address", "Mobile Phone", "Home Phone", "Business Phone",
        "Home Street", "Birthday", "Notes"
    ]
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for c in contacts:
            phones = c.get("phones", [])
            emails = c.get("emails", [])
            addrs = c.get("addresses", [])
            writer.writerow({
                "First Name": c.get("first_name", ""),
                "Middle Name": c.get("middle_name", ""),
                "Last Name": c.get("last_name", ""),
                "Title": c.get("title", ""),
                "Company": c.get("organization", ""),
                "E-mail Address": strip_label_suffix(emails[0]) if emails else "",
                "Mobile Phone": strip_label_suffix(phones[0]) if len(phones) > 0 else "",
                "Home Phone": strip_label_suffix(phones[1]) if len(phones) > 1 else "",
                "Business Phone": strip_label_suffix(phones[2]) if len(phones) > 2 else "",
                "Home Street": strip_label_suffix(addrs[0]) if addrs else "",
                "Birthday": normalize_date_string(c.get("birthday", "")),
                "Notes": c.get("note", ""),
            })


def ics_escape(value):
    value = str(value)
    value = value.replace("\\\\", "\\\\\\\\").replace(";", r"\\;").replace(",", r"\\,").replace("\\n", r"\\n")
    return value


def fold_ics_line(line, limit=75):
    encoded = line.encode("utf-8")
    if len(encoded) <= limit:
        return [line]
    out = []
    current = ""
    size = 0
    for ch in line:
        b = ch.encode("utf-8")
        if size + len(b) > limit:
            out.append(current)
            current = " " + ch
            size = len(current.encode("utf-8"))
        else:
            current += ch
            size += len(b)
    if current:
        out.append(current)
    return out


def write_google_birthdays_ics(contacts, out_path):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Contact Migrator//Birthday Export//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Imported Birthdays",
    ]
    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    counter = 1
    for c in contacts:
        bday = normalize_date_string(c.get("birthday", ""))
        if not re.match(r"^\\d{4}-\\d{2}-\\d{2}$", bday):
            continue
        name = clean_text(c.get("full_name", "")) or f"Unknown Contact {counter}"
        dtstart = bday.replace("-", "")
        summary = f"BD-{name} {bday}"
        desc = f"Birthday: {bday}"

        uid = f"birthday-{counter}-{re.sub(r'[^a-z0-9]+','-',name.lower()).strip('-')}@local"
        counter += 1

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_utc}",
            f"SUMMARY:{ics_escape(summary)}",
            f"DTSTART;VALUE=DATE:{dtstart}",
            "RRULE:FREQ=YEARLY",
            f"DESCRIPTION:{ics_escape(desc)}",
            "TRANSP:TRANSPARENT",
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            "DESCRIPTION:Birthday reminder",
            "TRIGGER:-PT10M",
            "END:VALARM",
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            "DESCRIPTION:Birthday reminder",
            "TRIGGER:-P1D",
            "END:VALARM",
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            "DESCRIPTION:Birthday reminder",
            "TRIGGER:-P10D",
            "END:VALARM",
            "END:VEVENT",
        ]

    lines.append("END:VCALENDAR")

    with open(out_path, "w", encoding="utf-8", newline="\\r\\n") as f:
        for line in lines:
            for folded in fold_ics_line(line):
                f.write(folded + "\\r\\n")


def process_files(project_dir, iphone_contacts=None, iphone_calendar=None, google_calendar=None, outlook_contacts=None):
    project_dir, input_dir, results_dir = ensure_workspace(project_dir)

    copied = {
        "iphone_contacts": copy_if_present(iphone_contacts, input_dir),
        "iphone_calendar": copy_if_present(iphone_calendar, input_dir),
        "google_calendar": copy_if_present(google_calendar, input_dir),
        "outlook_contacts": copy_if_present(outlook_contacts, input_dir),
    }

    contacts = []

    if copied["iphone_contacts"]:
        suffix = copied["iphone_contacts"].suffix.lower()
        if suffix == ".vcf":
            contacts.extend(parse_vcf(copied["iphone_contacts"]))
        elif suffix == ".csv":
            contacts.extend(parse_contacts_csv(copied["iphone_contacts"]))

    if copied["outlook_contacts"] and copied["outlook_contacts"].suffix.lower() == ".csv":
        contacts.extend(parse_contacts_csv(copied["outlook_contacts"]))

    if copied["iphone_calendar"] and copied["iphone_calendar"].suffix.lower() == ".ics":
        contacts.extend(parse_birthdays_ics(copied["iphone_calendar"]))

    if copied["google_calendar"] and copied["google_calendar"].suffix.lower() == ".ics":
        contacts.extend(parse_birthdays_ics(copied["google_calendar"]))

    contacts = dedupe_contacts(contacts)

    combined_csv = results_dir / "combined_contacts.csv"
    combined_xlsx = results_dir / "combined_contacts.xlsx"
    iphone_vcf = results_dir / "iphone_contacts_import.vcf"
    outlook_csv = results_dir / "outlook_contacts_import.csv"
    google_ics = results_dir / "google_birthdays_import.ics"
    summary_txt = results_dir / "summary.txt"

    write_combined_csv(contacts, combined_csv)
    xlsx_ok = write_xlsx_if_possible(combined_csv, combined_xlsx)
    write_iphone_vcf(contacts, iphone_vcf)
    write_outlook_csv(contacts, outlook_csv)
    write_google_birthdays_ics(contacts, google_ics)

    birthday_count = sum(1 for c in contacts if normalize_date_string(c.get("birthday", "")))
    with open(summary_txt, "w", encoding="utf-8") as f:
        f.write("Contact Migrator Summary\\n")
        f.write("========================\\n\\n")
        f.write(f"Project folder: {project_dir}\\n")
        f.write(f"Combined contacts: {len(contacts)}\\n")
        f.write(f"Contacts with birthdays: {birthday_count}\\n")
        f.write(f"Excel created: {'yes' if xlsx_ok else 'no'}\\n\\n")
        f.write("Inputs copied into input/:\\n")
        for k, v in copied.items():
            f.write(f"- {k}: {v if v else 'not provided'}\\n")
        f.write("\\nOutputs in results/:\\n")
        f.write(f"- {combined_csv}\\n")
        if xlsx_ok:
            f.write(f"- {combined_xlsx}\\n")
        f.write(f"- {iphone_vcf}\\n")
        f.write(f"- {outlook_csv}\\n")
        f.write(f"- {google_ics}\\n")
        f.write(f"- {summary_txt}\\n")

    return {
        "project_dir": str(project_dir),
        "input_dir": str(input_dir),
        "results_dir": str(results_dir),
        "combined_contacts": len(contacts),
        "birthday_count": birthday_count,
        "xlsx_created": xlsx_ok,
        "outputs": {
            "combined_csv": str(combined_csv),
            "combined_xlsx": str(combined_xlsx) if xlsx_ok else "",
            "iphone_vcf": str(iphone_vcf),
            "outlook_csv": str(outlook_csv),
            "google_ics": str(google_ics),
            "summary_txt": str(summary_txt),
        },
    }
"""

GUI_PY = r"""
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from .core import process_files


MAC_EXPORT_TEXT = \"\"\"Mac / iPhone export instructions

Apple Contacts:
- Open Contacts app
- Select the contacts you want
- File -> Export -> Export vCard
- Save the .vcf file

Apple / iPhone birthdays or calendar:
- Export a calendar as .ics if available
- If not available, skip this field

Google Calendar:
- Settings -> Import & export -> Export
- Unzip if needed
- Choose the correct .ics file
\"\"\"

WINDOWS_EXPORT_TEXT = \"\"\"Windows / Outlook export instructions

Outlook Contacts:
- Open Outlook
- Go to People / Contacts
- Export contacts as CSV

Google Calendar:
- Settings -> Import & export -> Export
- Unzip if needed
- Choose the correct .ics file

Apple / iPhone files:
- If already exported on another device, just browse and select them
\"\"\"


class App(tk.Tk):
    def __init__(self, platform_mode="mac"):
        super().__init__()
        self.title("Contact Migrator")
        self.geometry("840x700")

        self.platform_mode = platform_mode
        self.vars = {
            "project_dir": tk.StringVar(),
            "iphone_contacts": tk.StringVar(),
            "iphone_calendar": tk.StringVar(),
            "google_calendar": tk.StringVar(),
            "outlook_contacts": tk.StringVar(),
        }

        self._build_ui()

    def _build_ui(self):
        title = tk.Label(self, text="Contact Migrator", font=("Arial", 18, "bold"))
        title.pack(pady=(12, 6))

        instructions = MAC_EXPORT_TEXT if self.platform_mode == "mac" else WINDOWS_EXPORT_TEXT
        text = tk.Text(self, height=13, wrap="word")
        text.insert("1.0", instructions)
        text.configure(state="disabled")
        text.pack(fill="x", padx=14, pady=(0, 10))

        frame = tk.Frame(self)
        frame.pack(fill="both", expand=True, padx=14, pady=8)

        self._file_row(frame, 0, "Project folder", "project_dir", is_dir=True)
        self._file_row(frame, 1, "iPhone / Apple Contacts (.vcf or .csv)", "iphone_contacts")
        self._file_row(frame, 2, "iPhone / Apple Calendar birthdays (.ics)", "iphone_calendar")
        self._file_row(frame, 3, "Google Calendar export (.ics)", "google_calendar")
        self._file_row(frame, 4, "Outlook Contacts (.csv)", "outlook_contacts")

        note = tk.Label(
            frame,
            text="You can leave inputs blank. The app will process whatever you provide.",
            anchor="w",
            justify="left",
        )
        note.grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 6))

        run_btn = tk.Button(frame, text="Run", command=self.run_process, height=2)
        run_btn.grid(row=6, column=0, sticky="w", pady=(12, 8))

        quit_btn = tk.Button(frame, text="Quit", command=self.destroy, height=2)
        quit_btn.grid(row=6, column=1, sticky="w", pady=(12, 8))

    def _file_row(self, parent, row, label_text, key, is_dir=False):
        label = tk.Label(parent, text=label_text, anchor="w", justify="left")
        label.grid(row=row, column=0, sticky="w", pady=6)

        entry = tk.Entry(parent, textvariable=self.vars[key], width=70)
        entry.grid(row=row, column=1, sticky="we", padx=8, pady=6)

        def browse():
            if is_dir:
                path = filedialog.askdirectory(title=label_text)
            else:
                path = filedialog.askopenfilename(title=label_text)
            if path:
                self.vars[key].set(path)

        btn = tk.Button(parent, text="Browse", command=browse)
        btn.grid(row=row, column=2, sticky="w", pady=6)

        parent.grid_columnconfigure(1, weight=1)

    def run_process(self):
        project_dir = self.vars["project_dir"].get().strip()
        if not project_dir:
            messagebox.showerror("Missing project folder", "Please choose a project folder first.")
            return

        try:
            result = process_files(
                project_dir=project_dir,
                iphone_contacts=self.vars["iphone_contacts"].get().strip() or None,
                iphone_calendar=self.vars["iphone_calendar"].get().strip() or None,
                google_calendar=self.vars["google_calendar"].get().strip() or None,
                outlook_contacts=self.vars["outlook_contacts"].get().strip() or None,
            )
        except Exception as e:
            messagebox.showerror("Processing error", str(e))
            return

        outputs = result["outputs"]
        msg = (
            f"Done.\\n\\n"
            f"Combined contacts: {result['combined_contacts']}\\n"
            f"Contacts with birthdays: {result['birthday_count']}\\n"
            f"Excel created: {'yes' if result['xlsx_created'] else 'no'}\\n\\n"
            f"Results folder:\\n{result['results_dir']}\\n\\n"
            f"Main outputs:\\n"
            f"- {outputs['combined_csv']}\\n"
            f"- {outputs['iphone_vcf']}\\n"
            f"- {outputs['outlook_csv']}\\n"
            f"- {outputs['google_ics']}\\n"
        )
        messagebox.showinfo("Success", msg)


def main():
    platform_mode = "mac"
    if len(sys.argv) > 1:
        platform_mode = sys.argv[1].lower().strip()
        if platform_mode not in {"mac", "windows"}:
            platform_mode = "mac"

    app = App(platform_mode=platform_mode)
    app.mainloop()


if __name__ == "__main__":
    main()
"""

RUN_MAC = """#!/bin/bash
cd "$(dirname "$0")/.."
python3 -m contact_migrator.gui mac
"""

RUN_WINDOWS = """@echo off
cd /d "%~dp0\\.."
py -m contact_migrator.gui windows
"""


def write_file(path: Path, content: str, executable: bool = False):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")
    if executable:
        try:
            path.chmod(0o755)
        except Exception:
            pass


def write_project_files(base_dir: Path):
    files = {
        base_dir / "README.md": README_MD,
        base_dir / "requirements.txt": REQUIREMENTS_TXT,
        base_dir / ".gitignore": GITIGNORE,
        base_dir / "index.md": INDEX_MD,
        base_dir / "contact_migrator" / "__init__.py": INIT_PY,
        base_dir / "contact_migrator" / "core.py": CORE_PY.strip() + "\n",
        base_dir / "contact_migrator" / "gui.py": GUI_PY.strip() + "\n",
        base_dir / "scripts" / "run_mac.command": RUN_MAC,
        base_dir / "scripts" / "run_windows.bat": RUN_WINDOWS,
    }

    for path, content in files.items():
        write_file(path, content, executable=path.name.endswith(".command"))
    print(f"Wrote project files into: {base_dir}")


def print_explanation(base_dir: Path):
    print("Contact Migrator bootstrap")
    print("==========================")
    print("")
    print("This script can:")
    print("- write the full project files into the current repository")
    print("- run the GUI on macOS or Windows after the files exist")
    print("")
    print(f"Current folder: {base_dir}")
    print("")
    print("Suggested sequence:")
    print("1. python3 bootstrap_contact_migrator.py --write")
    print("2. python3 -m venv .venv")
    print("3. source .venv/bin/activate    # macOS")
    print("4. pip install -r requirements.txt")
    print("5. python3 bootstrap_contact_migrator.py --run-mac")
    print("")
    print("Windows:")
    print("1. py bootstrap_contact_migrator.py --write")
    print("2. py -m venv .venv")
    print(r"3. .venv\Scripts\Activate.ps1")
    print("4. pip install -r requirements.txt")
    print("5. py bootstrap_contact_migrator.py --run-windows")
    print("")


def run_gui(mode: str, base_dir: Path):
    core_file = base_dir / "contact_migrator" / "gui.py"
    if not core_file.exists():
        print("Project files do not exist yet.")
        print("Run with --write first.")
        sys.exit(1)

    if mode == "mac":
        cmd = [sys.executable, "-m", "contact_migrator.gui", "mac"]
    else:
        cmd = [sys.executable, "-m", "contact_migrator.gui", "windows"]

    subprocess.run(cmd, cwd=str(base_dir), check=False)


def main():
    parser = argparse.ArgumentParser(description="Write and run the Contact Migrator project.")
    parser.add_argument("--write", action="store_true", help="Write the full project files into the current folder.")
    parser.add_argument("--run-mac", action="store_true", help="Run the GUI in macOS mode.")
    parser.add_argument("--run-windows", action="store_true", help="Run the GUI in Windows mode.")
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    print_explanation(base_dir)

    if args.write:
        write_project_files(base_dir)

    if args.run_mac:
        run_gui("mac", base_dir)

    if args.run_windows:
        run_gui("windows", base_dir)

    if not (args.write or args.run_mac or args.run_windows):
        print("Nothing selected.")
        print("Use --write, --run-mac, or --run-windows.")


if __name__ == "__main__":
    main()
