# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Convert the Dalton jaaragenda spreadsheet into one .ics calendar per column.

The columns to convert, and the name to use for each, come from --column, from
columns.toml, or are asked interactively and then saved to columns.toml. The
name becomes the event prefix and the file name.

Each cell line becomes an all-day event. Identical entries on consecutive
days are merged into one multi-day event, and vacations ("start ...vakantie"
up to "laatste dag ...") become a single event. UIDs are derived from the
content, so re-running on a newer version of the spreadsheet updates a
subscribed calendar instead of duplicating events.
"""

import argparse
import datetime as dt
import hashlib
import os
import re
import sys
import tomllib
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl

# Stable link on the school website; it redirects to the current .xlsx.
SOURCE_URL = "https://haarlemmermeerlyceum.nl/jaaragenda-dalton"
# The spreadsheet itself is in Dutch: month names and header names.
MONTHS = {m: i for i, m in enumerate(
    "januari februari maart april mei juni juli augustus september oktober november december".split(), 1)}
DATE_HEADER = "datum"
SKIP_HEADERS = {"week", "dag", DATE_HEADER}


# --- reading the spreadsheet -------------------------------------------------

def find_columns(ws):
    """Return (date column index, {header: column index}) from the header row."""
    headers = {c.column: str(c.value).strip() for c in ws[1] if c.value not in (None, "")}
    date_col = next((col for col, h in headers.items() if h.lower() == DATE_HEADER), None)
    if date_col is None:
        sys.exit("No 'Datum' column found in the first row of the sheet.")
    columns = {}
    for col, header in sorted(headers.items()):
        if header.lower() not in SKIP_HEADERS and header not in columns:
            columns[header] = col
    return date_col, columns


def clean(text):
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r" ,", ",", text)
    return text[:1].upper() + text[1:]


def read_column(ws, date_col, col):
    """Return a sorted list of (date, line) for every non-empty line in the column."""
    items = []
    for r in range(2, ws.max_row + 1):
        day = ws.cell(r, date_col).value
        val = ws.cell(r, col).value
        if not isinstance(day, dt.datetime) or not isinstance(val, str):
            continue
        for line in val.splitlines():
            if line.strip():
                items.append((day.date(), clean(line)))
    return sorted(items)


def school_year(ws, date_col):
    days = [v for (v,) in ws.iter_rows(min_row=2, min_col=date_col, max_col=date_col, values_only=True)
            if isinstance(v, dt.datetime)]
    return f"{min(days).year}-{max(days).year}" if days else ""


# --- turning cells into events ----------------------------------------------

def extract_vacations(items):
    """Turn start/end vacation markers into (start, end, title) ranges."""
    vacations, rest, open_vac = [], [], None
    for day, text in items:
        low = text.lower()
        m = re.match(r"(?:start )?(\w*vakantie)\b(.*)", low)
        if m and not low.startswith("laatste"):
            name = m.group(1).capitalize()
            until = re.search(r"t/m (\d+) (\w+)", m.group(2))
            if until and until.group(2) in MONTHS:  # e.g. "start zomervakantie t/m 22 augustus"
                month = MONTHS[until.group(2)]
                year = day.year + (1 if month < day.month else 0)
                vacations.append((day, dt.date(year, month, int(until.group(1))), name))
            else:
                open_vac = (day, name)
            continue
        if open_vac and low.startswith("laatste") and "vakantie" in low:
            vacations.append((open_vac[0], day, open_vac[1]))
            open_vac = None
            continue
        rest.append((day, text))
    if open_vac:  # no end marker found: keep it as a single-day event
        vacations.append((open_vac[0], open_vac[0], f"Start {open_vac[1].lower()}"))
    return vacations, rest


def merge_consecutive(items):
    """Merge identical (case-insensitive) entries on consecutive days into ranges."""
    ranges = []
    open_by_text = {}
    for day, text in items:
        key = text.lower()
        prev = open_by_text.get(key)
        if prev is not None and ranges[prev][1] == day - dt.timedelta(days=1):
            ranges[prev] = (ranges[prev][0], day, ranges[prev][2])
        else:
            open_by_text[key] = len(ranges)
            ranges.append((day, day, text))
    return ranges


# --- writing iCalendar -------------------------------------------------------

def escape(text):
    return text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(line):
    """Fold content lines at 75 octets as required by RFC 5545."""
    raw = line.encode()
    out = []
    while len(raw) > 75:
        cut = 75 if not out else 74
        while (raw[cut] & 0xC0) == 0x80:  # do not split a UTF-8 sequence
            cut -= 1
        out.append(raw[:cut].decode())
        raw = raw[cut:]
    out.append(raw.decode())
    return "\r\n ".join(out)


def build_ics(cal_name, column, prefix, ranges, source, stamp):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//hlml-dalton-jaaragenda//NL",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{escape(cal_name)}",
        f"NAME:{escape(cal_name)}",
        "X-WR-TIMEZONE:Europe/Amsterdam",
        "REFRESH-INTERVAL;VALUE=DURATION:PT12H",
        "X-PUBLISHED-TTL:PT12H",
    ]
    for start, end, title in ranges:
        uid = hashlib.sha1(f"{column}|{start}|{title.lower()}".encode()).hexdigest()
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}@hlml-dalton-jaaragenda",
            f"DTSTAMP:{stamp}",
            f"DTSTART;VALUE=DATE:{start:%Y%m%d}",
            f"DTEND;VALUE=DATE:{end + dt.timedelta(days=1):%Y%m%d}",  # exclusive
            f"SUMMARY:{escape(f'{prefix}: {title}')}",
            f"DESCRIPTION:{escape(f'{column}. Source: {source}')}",
            "TRANSP:TRANSPARENT",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "".join(fold(line) + "\r\n" for line in lines)


def slug(text):
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "calendar"


# --- choosing columns --------------------------------------------------------

def ask(prompt):
    try:
        return input(prompt).strip()
    except EOFError:
        sys.exit("\nAborted: no input. Use --column or columns.toml to run without questions.")


def ask_columns(columns):
    """Interactively pick columns and a name for each; returns [(column, name)]."""
    headers = list(columns)
    print("Columns in this jaaragenda:\n")
    for i, header in enumerate(headers, 1):
        print(f"  {i:2}. {header}")
    print()
    while True:
        answer = ask("Which columns do you want? Enter numbers separated by commas (e.g. 1,2,5): ")
        try:
            picked = [headers[int(n) - 1] for n in re.split(r"[,\s]+", answer) if n]
        except (ValueError, IndexError):
            picked = []
        if picked:
            break
        print(f"Invalid choice. Pick numbers between 1 and {len(headers)}.")

    print("\nGive each column a name, for example your child's name. It is put in front of")
    print("every event (\"Sam: Schoolfotograaf\") and becomes the file name.\n")
    chosen = []
    for header in dict.fromkeys(picked):
        default = "Dalton" if header.lower() == "algemeen" else header
        name = ask(f"  Name for '{header}' [{default}]: ") or default
        chosen.append((header, name))
    return chosen


def toml_str(text):
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def load_config(path):
    with path.open("rb") as f:
        columns = tomllib.load(f).get("columns", {})
    if not columns:
        sys.exit(f"{path} has no [columns] table.")
    return [(str(column), str(name)) for column, name in columns.items()]


def save_config(path, chosen):
    lines = [
        "# Columns from the jaaragenda to convert, and the name to use for each.",
        "# The name is put in front of every event and becomes the .ics file name.",
        "[columns]",
        *(f"{toml_str(column)} = {toml_str(name)}" for column, name in chosen),
    ]
    path.write_text("\n".join(lines) + "\n")


def parse_column(value):
    column, sep, name = value.partition("=")
    if not sep or not column.strip() or not name.strip():
        raise argparse.ArgumentTypeError(f"expected 'COLUMN=NAME', got '{value}'")
    return column.strip(), name.strip()


# --- the source spreadsheet --------------------------------------------------

def xlsx_files(folder):
    return [p for p in sorted(folder.glob("*.xlsx")) if not p.name.startswith("~$")]  # skip Excel lock files


def fetch(url, folder):
    """Download the jaaragenda into folder, replacing any older version there."""
    if not url.startswith(("https://", "http://")):
        sys.exit(f"--fetch expects a URL, not '{url}'. Pass a local spreadsheet without --fetch.")
    request = urllib.request.Request(url, headers={"User-Agent": "hlml-dalton-jaaragenda"})
    with urllib.request.urlopen(request, timeout=60) as response:
        final_url = response.geturl()
        data = response.read()
    name = Path(urllib.parse.unquote(urllib.parse.urlparse(final_url).path)).name
    if not name.lower().endswith(".xlsx") or not data.startswith(b"PK"):
        sys.exit(f"{url} did not return a spreadsheet but {final_url}")

    folder.mkdir(parents=True, exist_ok=True)
    target = folder / name
    if target.exists() and target.read_bytes() == data:
        print(f"No new version: {name} is unchanged.")
    else:
        target.write_bytes(data)
        print(f"Fetched new version: {name}")
    for old in xlsx_files(folder):
        if old != target:
            old.unlink()
            print(f"Removed old version: {old.name}")
    return target


def default_xlsx(folder):
    found = xlsx_files(folder)
    if len(found) != 1:
        listed = "".join(f"\n  {p}" for p in found) or " none"
        sys.exit(f"Pass the spreadsheet explicitly; {os.path.relpath(folder)}/ does not hold exactly one:{listed}")
    return found[0]


def main():
    here = Path(__file__).parent
    parser = argparse.ArgumentParser(
        description="Convert columns of the Dalton jaaragenda (.xlsx) into calendar files (.ics).")
    parser.add_argument("xlsx", type=Path, nargs="?",
                        help="the jaaragenda spreadsheet (default: the only .xlsx file in source/)")
    parser.add_argument("--fetch", nargs="?", const=SOURCE_URL, metavar="URL",
                        help=f"first download the newest jaaragenda into source/, replacing the old one "
                             f"(default URL: {SOURCE_URL})")
    parser.add_argument("--column", action="append", type=parse_column, metavar="COLUMN=NAME",
                        help="column and the name to prefix its events with, e.g. 'Leerjaar 1=Sam' "
                             "(repeatable); overrides columns.toml")
    parser.add_argument("--config", type=Path, default=here / "columns.toml",
                        help="file with the chosen columns (default: columns.toml); "
                             "if it does not exist you are asked and the answers are saved there")
    parser.add_argument("--output", type=Path, default=here / "calendars",
                        help="directory for the .ics files (default: calendars/)")
    args = parser.parse_args()
    if args.fetch and args.xlsx:
        parser.error("use either --fetch or a spreadsheet, not both")
    if args.fetch:
        xlsx = fetch(args.fetch, here / "source")
    else:
        xlsx = args.xlsx or default_xlsx(here / "source")

    wb = openpyxl.load_workbook(xlsx)
    ws = wb.worksheets[0]
    date_col, columns = find_columns(ws)

    if args.column:
        chosen = args.column
    elif args.config.exists():
        chosen = load_config(args.config)
        print(f"Columns from {os.path.relpath(args.config)}")
    else:
        chosen = ask_columns(columns)
        save_config(args.config, chosen)
        print(f"\nSaved your choice to {os.path.relpath(args.config)}; edit or delete it to choose again.")

    unknown = [c for c, _ in chosen if c not in columns]
    if unknown:
        sys.exit(f"Unknown column(s): {', '.join(unknown)}. Available: {', '.join(columns)}")

    # Derive DTSTAMP from the spreadsheet, so unchanged input gives identical output.
    modified = wb.properties.modified or dt.datetime(2000, 1, 1)
    stamp = modified.replace(tzinfo=None).strftime("%Y%m%dT%H%M%SZ")
    year = school_year(ws, date_col)
    args.output.mkdir(parents=True, exist_ok=True)
    print()
    used = set()
    for column, name in chosen:
        fname = slug(name)
        while fname in used:
            fname += "-2"
        used.add(fname)
        vacations, rest = extract_vacations(read_column(ws, date_col, columns[column]))
        ranges = sorted(vacations + merge_consecutive(rest))
        cal_name = f"{name} - {column} ({year})"
        path = args.output / f"{fname}.ics"
        path.write_text(build_ics(cal_name, column, name, ranges, xlsx.name, stamp), newline="")
        print(f"{os.path.relpath(path)}: {len(ranges)} events for '{column}'")


if __name__ == "__main__":
    main()
