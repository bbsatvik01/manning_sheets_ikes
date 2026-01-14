"""
Manning Chart Web Application
============================

This script provides a simple web interface for generating Manning Chart
workbooks from staff scheduling spreadsheets.  When run (or packaged
as a Windows executable with PyInstaller), it starts a local web
server on `http://127.0.0.1:5000` and opens a page where you can
upload schedule files, trigger processing, and download or view the
resulting Manning Charts.  Uploaded schedule files are stored in
``input_my_staff_schedule``, generated charts in ``manning_sheets``,
and a log file is written to ``logs``.

Requirements
------------

* ``flask`` for the web server.  Install it via ``pip install flask``.
* ``openpyxl`` for reading and writing Excel files.

Packaging as an executable
-------------------------

You can package this script and its dependencies into a single
executable using PyInstaller.  First install PyInstaller::

    pip install pyinstaller

Then create the executable::

    pyinstaller --onefile manning_web_app.py

The resulting ``manning_web_app.exe`` (found in the ``dist``
directory) can be run on Windows without requiring Python to be
pre‑installed.  Double‑clicking the executable will start the local
server and open the upload page in your default web browser.

"""

import os
import re
import sys
import threading
import logging
import webbrowser
from datetime import datetime
from typing import Any, Dict, List, Optional

from flask import (
    Flask,
    request,
    redirect,
    url_for,
    render_template_string,
    send_from_directory,
    abort,
    flash,
    get_flashed_messages,
)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
except ImportError as exc:
    raise SystemExit(
        "Required dependency openpyxl is missing. Install with 'pip install openpyxl'."
    )


###############################
# Configuration and utilities #
###############################

# Determine base directory for read/write data and locate bundled resources
if getattr(sys, "frozen", False):
    EXEC_DIR = os.path.dirname(sys.executable)
    RESOURCE_DIR = getattr(sys, "_MEIPASS", EXEC_DIR)
else:
    EXEC_DIR = os.path.dirname(os.path.abspath(__file__))
    RESOURCE_DIR = EXEC_DIR

BASE_DIR = EXEC_DIR
INPUT_DIR = os.path.join(BASE_DIR, "input_my_staff_schedule")
OUTPUT_DIR = os.path.join(BASE_DIR, "manning_sheets")
LOG_DIR = os.path.join(BASE_DIR, "logs")
STATIC_ROOT = os.path.join(RESOURCE_DIR, "static")
if not os.path.isdir(STATIC_ROOT):
    STATIC_ROOT = os.path.join(BASE_DIR, "static")

# Track outputs generated during this runtime (latest batch only)
CURRENT_OUTPUTS: List[str] = []

# Ensure directories exist
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# Set up logging
log_filename = os.path.join(LOG_DIR, "manning_app.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)


def parse_time(time_str: str) -> Optional[float]:
    """Convert a 12‑hour time string into a floating point hour.

    Example: "06:00 AM" -> 6.0, "02:30 PM" -> 14.5.
    Returns None if parsing fails.
    """
    time_str = time_str.strip()
    match = re.match(r"(\d{1,2}):(\d{2})\s*(AM|PM)", time_str, re.IGNORECASE)
    if not match:
        return None
    hour, minute, meridiem = int(match.group(1)), int(match.group(2)), match.group(3).upper()
    if meridiem == "PM" and hour != 12:
        hour += 12
    if meridiem == "AM" and hour == 12:
        hour = 0
    return hour + minute / 60.0


def get_category(role: str) -> Optional[str]:
    """Map a job role to a Manning Chart category (station)."""
    if not role:
        return None
    role_clean = role.strip().replace("\n", "").strip()
    role_upper = role_clean.upper()

    # Student roles
    if role_upper.startswith("STUDENT "):
        if "HOMESTYLE" in role_upper:
            return "HOMESTYLE ROOTED"
        if "DWO" in role_upper:
            return "DELICIOUS WITHOUT"
        if "HOMESLICE" in role_upper:
            return "HOMESLICE"
        if "UNITED TABLE" in role_upper:
            return "UNITED TABLE"
        if "FLIPS" in role_upper:
            return "FLIPS"
        if "PIZZA/PASTA" in role_upper:
            return "FLOUR SAUCE"
        if "DESSERTS" in role_upper:
            return "SWEET SHOPPE"
        if "FOH" in role_upper:
            return "BEVERAGES"
        if "SALAD BAR" in role_upper:
            return "GARDEN SOCIAL & NOOK"
        if "UTILITY" in role_upper:
            return "UTILITY"
        if "SUPERVISOR" in role_upper:
            return "SUPERVISOR"
        return None

    # Full‑time roles
    if role_upper.startswith("PRODUCTION COOK"):
        return "HOMESTYLE ROOTED"
    if role_upper.startswith("DWO COOK"):
        return "DELICIOUS WITHOUT"
    if role_upper.startswith("FLIPS COOK"):
        return "FLIPS"
    if role_upper.startswith("PIZZA COOK"):
        return "FLOUR SAUCE"
    if role_upper.startswith("UNITED TABLE COOK"):
        return "UNITED TABLE"
    if role_upper.startswith("DELI COOK"):
        return "HOMESLICE"
    if role_upper.startswith("COLD PREP COOK"):
        return "GARDEN SOCIAL & NOOK"
    if role_upper.startswith("UTILITY DISHROOM") or role_upper.startswith("UTILITY POTS") or role_upper.startswith("UTILITY FOH"):
        return "UTILITY"
    if role_upper.startswith("CASHIER"):
        return "CASHIER"
    if "SUPERVISOR" in role_upper:
        return "SUPERVISOR"
    return None


def process_schedule_file(file_path: str, output_dir: str) -> List[str]:
    """Process a single schedule Excel file and generate Manning Charts.

    Returns a list of output file paths created.
    """
    outputs: List[str] = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as exc:
        logging.error(f"Error opening '{file_path}': {exc}")
        return outputs

    ws = wb.active
    header_cell_value = ws['A1'].value or ""
    year_matches = re.findall(r"\d{4}", str(header_cell_value))
    year = year_matches[-1] if year_matches else ""

    header_row = [cell for cell in ws.iter_rows(min_row=2, max_row=2, values_only=True)][0]
    date_columns: List[int] = []
    date_labels: List[str] = []
    for idx, cell in enumerate(header_row):
        if idx == 0 or not cell:
            continue
        m = re.search(r"(\d{2}/\d{2})", str(cell))
        if m:
            date_columns.append(idx)
            label = m.group(1)
            if year and year not in label:
                label = f"{label}/{year}"
            date_labels.append(label)
    if not date_columns:
        logging.warning(f"No valid date columns in '{file_path}'.")
        return outputs

    # Shift definitions: inclusive boundaries
    shifts = [
        {'name': '6am-2pm', 'meal_periods': 'B BR', 'lower': 6, 'upper': 14},  # 6 ≤ start < 14
        {'name': '2pm-11pm', 'meal_periods': 'D', 'lower': 14, 'upper': 22},    # 14 ≤ start < 22
        {'name': '10pm-6am', 'meal_periods': 'OVNT', 'lower': 22, 'upper': 24}, # start ≥22 or start <6
    ]

    # Chart row layout
    row_groups = [
        ['CASHIER', 'GARDEN SOCIAL & NOOK', 'LA CUCINA'],
        ['FLOUR SAUCE', 'DELICIOUS WITHOUT', 'SWEET SHOPPE'],
        ['UNITED TABLE', 'HOMESLICE', 'HOMESTYLE ROOTED'],
        ['FLIPS', 'CULINARY', 'UTILITY'],
        ['BEVERAGES', 'TABLE BUSSER', 'SUPERVISOR'],
    ]

    generation_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # For each date, produce a workbook
    for file_counter, (col_idx, date_label) in enumerate(zip(date_columns, date_labels), start=1):
        try:
            parsed_date = datetime.strptime(date_label, "%m/%d/%Y")
            weekday_label = f" ({parsed_date.strftime('%A')})"
            date_part = parsed_date.strftime("%a_%d_%b")
        except ValueError:
            parsed_date = None
            weekday_label = ""
            date_part = date_label.replace('/', '-').replace('-', '_')
        shift_data = [
            {category: [] for group in row_groups for category in group}
            for _ in shifts
        ]
        for row in ws.iter_rows(min_row=3, values_only=True):
            role = row[0]
            cell_val = row[col_idx]
            if not cell_val:
                continue
            category = get_category(str(role))
            if not category:
                continue
            cell_str = str(cell_val).strip()
            assignments = [s for s in re.split(r"\n{2,}", cell_str) if s.strip()]
            i = 0
            while i < len(assignments):
                if i + 1 < len(assignments):
                    name = assignments[i].strip()
                    time_range = assignments[i + 1].strip()
                    i += 2
                else:
                    break
                m = re.match(
                    r"(\d{1,2}:\d{2}\s*\w{2})\s*-\s*(\d{1,2}:\d{2}\s*\w{2})",
                    time_range,
                )
                if not m:
                    continue
                start_time = parse_time(m.group(1))
                if start_time is None:
                    continue
                if 6 <= start_time < 14:
                    shift_index = 0
                elif 14 <= start_time < 22:
                    shift_index = 1
                else:
                    shift_index = 2
                entry = f"{name}\n{m.group(1)} - {m.group(2)}"
                shift_data[shift_index][category].append(entry)

        # Create output workbook
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)
        for idx_shift, shift_info in enumerate(shifts):
            sheet = out_wb.create_sheet(shift_info['name'])
            sheet.merge_cells('A1:C1')
            sheet['A1'] = 'MANNING CHART'
            sheet['A1'].font = Font(size=14, bold=True)
            sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
            sheet.merge_cells('A2:C2')
            sheet['A2'] = f"Date: {date_label}{weekday_label}    Meal Periods: {shift_info['meal_periods']}    MOD:"
            sheet['A2'].alignment = Alignment(horizontal='left', vertical='center')
            for col_letter in ['A', 'B', 'C']:
                sheet.column_dimensions[col_letter].width = 30
            thin = Side(border_style='thin', color='000000')
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            row_ptr = 3
            for group in row_groups:
                for col, label in enumerate(group, start=1):
                    hcell = sheet.cell(row=row_ptr, column=col)
                    hcell.value = label
                    hcell.font = Font(bold=True)
                    hcell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    hcell.border = border
                data_row = row_ptr + 1
                for col, label in enumerate(group, start=1):
                    dcell = sheet.cell(row=data_row, column=col)
                    items = shift_data[idx_shift].get(label, [])
                    dcell.value = '\n\n'.join(items) if items else ''
                    dcell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                    dcell.border = border
                sheet.row_dimensions[row_ptr].height = 20
                sheet.row_dimensions[data_row].height = 60
                row_ptr += 2
        out_filename = f"{date_part}_Manning_sheet_{generation_stamp}.xlsx"
        out_path = os.path.join(output_dir, out_filename)
        try:
            out_wb.save(out_path)
            outputs.append(out_filename)
            logging.info(f"Generated '{out_filename}' from '{os.path.basename(file_path)}'.")
        except Exception as exc:
            logging.error(f"Error saving '{out_filename}': {exc}")
    return outputs


######################
# Flask web server   #
######################

app = Flask(__name__, static_folder=STATIC_ROOT, static_url_path="/static")
app.secret_key = os.environ.get("MANNING_APP_SECRET", "manning-standalone-secret")


BASE_CSS = """
/* UI styling lives in static/assets/css. Inline rules reserved for quick overrides. */
"""


def asset_urls() -> Dict[str, str]:
    """Return URLs for local CSS/JS assets served from /static."""
    return {
        "css_black": url_for("static", filename="assets/css/black-dashboard.min.css"),
        "css_custom": url_for("static", filename="assets/css/custom.css"),
        "css_icons": url_for("static", filename="assets/css/nucleo-icons.css"),
        "js_jquery": url_for("static", filename="assets/js/core/jquery.min.js"),
        "js_popper": url_for("static", filename="assets/js/core/popper.min.js"),
        "js_bootstrap": url_for("static", filename="assets/js/core/bootstrap.min.js"),
        "js_black": url_for("static", filename="assets/js/black-dashboard.min.js"),
    }


def parse_cell_assignments(cell_value: Optional[str]) -> List[Dict[str, str]]:
    """Split a cell value into individual staff assignments."""
    if not cell_value:
        return []
    text = str(cell_value).strip()
    if not text:
        return []
    blocks = [block.strip() for block in re.split(r"\n{2,}", text) if block.strip()]
    assignments: List[Dict[str, str]] = []
    for block in blocks:
        lines = [line.strip() for line in block.splitlines() if line.strip()]
        if not lines:
            continue
        name = lines[0]
        shift = " ".join(lines[1:]).strip() if len(lines) > 1 else ""
        assignments.append({"name": name, "time": shift})
    return assignments


def build_sheet_structure(ws: "openpyxl.worksheet.worksheet.Worksheet") -> Dict[str, Any]:
    """Create a structured representation of a worksheet's staffing data."""
    rows = list(ws.iter_rows(values_only=True))
    stations: List[Dict[str, Any]] = []
    excel_sections: List[Dict[str, List[str]]] = []
    row_idx = 2  # first two rows are header/meta rows
    while row_idx < len(rows):
        header_row = rows[row_idx]
        data_row = rows[row_idx + 1] if row_idx + 1 < len(rows) else None
        if header_row and any(header_row[:3]) and data_row is not None:
            headers: List[str] = []
            cells: List[str] = []
            for col_idx in range(3):
                header_val = header_row[col_idx] if col_idx < len(header_row) else None
                station_name = str(header_val).strip() if header_val else ""
                headers.append(station_name)
                if not station_name:
                    cells.append(str(data_row[col_idx]) if data_row and col_idx < len(data_row) and data_row[col_idx] else "")
                    continue
                cell_val = data_row[col_idx] if col_idx < len(data_row) else None
                assignments = parse_cell_assignments(cell_val)
                stations.append({"station": station_name, "entries": assignments})
                cells.append(str(cell_val) if cell_val else "")
            excel_sections.append({"headers": headers, "cells": cells})
            row_idx += 2
            continue
        row_idx += 1
    total_entries = sum(len(station["entries"]) for station in stations)
    return {"stations": stations, "total_entries": total_entries, "excel_sections": excel_sections}


def list_output_files() -> List[str]:
    """Return a sorted list of .xlsx files in the output directory."""
    files = [f for f in os.listdir(OUTPUT_DIR) if f.lower().endswith('.xlsx')]
    files.sort()
    return files


@app.route('/', methods=['GET'])
def index() -> str:
    """Render the upload form and list existing outputs."""
    view_mode = request.args.get("view", "current")
    show_history = view_mode == "history"
    files = list_output_files() if show_history else CURRENT_OUTPUTS.copy()
    total_generated = len(CURRENT_OUTPUTS)
    latest_file = CURRENT_OUTPUTS[-1] if CURRENT_OUTPUTS else None
    assets = asset_urls()
    flashes = get_flashed_messages(with_categories=True)
    return render_template_string(
        """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Manning Chart Generator</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap">
    <link rel="stylesheet" href="{{ css_black }}">
    <link rel="stylesheet" href="{{ css_custom }}">
    <link rel="stylesheet" href="{{ css_icons }}">
    <style>{{ base_css }}</style>
</head>
<body class="app-shell">
<main class="app-content single-column">
        {% if flashes %}
        <div class="toast-stack">
            {% for category, message in flashes %}
            <div class="toast {{ category }}">{{ message }}</div>
            {% endfor %}
        </div>
        {% endif %}
        <div class="page-heading">
            <h1>Manning Chart Generator</h1>
        </div>
        <div class="hero-row">
            <div class="hero-cell col-3 hero-desc">
                <p class="eyebrow">Automated coverage</p>
                <h2>Manning Chart Generator</h2>
                <p class="muted">Upload the latest <strong>.xlsx</strong> schedule. We keep the raw file in <code>input_my_staff_schedule</code> and save processed shifts inside <code>manning_sheets</code>.</p>
            </div>
            <div class="hero-cell col-6 hero-upload">
                <p class="eyebrow text-center">Upload schedule</p>
                <h2 class="text-center">Generate new Manning sheets</h2>
                <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data" class="upload-form centered">
                    <label for="file-input" class="muted">Drop or browse for a MyStaff Excel export (.xlsx)</label>
                    <input id="file-input" type="file" name="file" accept=".xlsx" required class="app-file-input">
                    <button type="submit" class="btn btn-primary btn-gradient">Upload &amp; Generate Charts</button>
                </form>
            </div>
            <div class="hero-cell col-3 hero-stats">
                <div class="metric">
                    <span>Total charts</span>
                    <strong>{{ total_generated }}</strong>
                </div>
                <div class="metric">
                    <span>Latest workbook</span>
                    <strong>{{ latest_file or "N/A" }}</strong>
                </div>
                <div class="hero-actions">
                    <a class="btn btn-info hero-log" href="{{ url_for('view_log') }}">View processing log</a>
                    <a class="btn btn-outline hero-log" href="{{ url_for('view_log') }}" target="_blank">Open log in new tab</a>
                </div>
            </div>
        </div>

        <section class="app-card">
            <div class="section-header">
                <div>
                    <p class="eyebrow">{{ "History" if show_history else "Session files" }}</p>
                    <h2>{{ "Generated Manning Sheets History" if show_history else "Your generated Manning Sheets" }}</h2>
                    <p class="muted note">
                        {% if show_history %}
                        All previously generated workbooks are listed here.
                        {% else %}
                        Files listed here were created in this session. All workbooks are still stored in <code>manning_sheets</code>.
                        {% endif %}
                    </p>
                </div>
                <div class="section-actions">
                    <a class="btn btn-sm btn-outline" href="{{ url_for('index', view='history') }}">History</a>
                    <a class="btn btn-sm btn-gradient" href="{{ url_for('index') }}">Current</a>
                </div>
            </div>
            {% if files %}
            <div class="table-scroll">
                <table class="file-table">
                    <thead>
                        <tr>
                            <th>Workbook</th>
                            <th style="width:280px;">Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for fname in files %}
                        <tr>
                            <td><span class="filename">{{ fname }}</span></td>
                            <td class="action-stack">
                                <a class="file-action inline" href="{{ url_for('download_file', filename=fname) }}">Download</a>
                                <a class="file-action inline" href="{{ url_for('view_file', filename=fname) }}">View online</a>
                            </td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
            {% else %}
            <p class="file-empty">
                {% if show_history %}
                No historical files found yet.
                {% else %}
                No Manning charts yet. Upload a schedule to kick things off.
                {% endif %}
            </p>
            {% endif %}
        </section>
    </main>
<script src="{{ js_jquery }}"></script>
<script src="{{ js_popper }}"></script>
<script src="{{ js_bootstrap }}"></script>
<script src="{{ js_black }}"></script>
</body>
</html>
        """,
        files=files,
        total_generated=total_generated,
        latest_file=latest_file,
        base_css=BASE_CSS,
        flashes=flashes,
        show_history=show_history,
        **assets,
    )


@app.route('/upload', methods=['POST'])
def upload() -> str:
    """Handle file upload, save to input directory, process it, and redirect."""
    global CURRENT_OUTPUTS
    if 'file' not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash("Please choose a file before uploading.", "error")
        return redirect(url_for('index'))
    if not file.filename.lower().endswith('.xlsx'):
        flash("Only .xlsx files are supported. Upload the MyStaff shift schedule exported in Task Wise view (.xlsx).", "error")
        return redirect(url_for('index'))
    # Save the uploaded file with a timestamp to avoid collisions
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_filename = f"{timestamp}_{os.path.basename(file.filename)}"
    input_path = os.path.join(INPUT_DIR, safe_filename)
    file.save(input_path)
    logging.info(f"Uploaded schedule saved to '{input_path}'.")
    # Process the uploaded schedule
    outputs = process_schedule_file(input_path, OUTPUT_DIR)
    if outputs:
        logging.info(f"Generated {len(outputs)} output file(s) from '{safe_filename}'.")
        CURRENT_OUTPUTS = outputs
        flash(f"Successfully generated {len(outputs)} chart(s).", "success")
    else:
        logging.warning(f"No output files generated from '{safe_filename}'.")
        flash("Unable to process that file. Upload the MyStaff shift schedule exported in Task Wise view (.xlsx) and try again.", "error")
    # Redirect to index to show new files
    return redirect(url_for('index'))


@app.route('/download/<path:filename>')
def download_file(filename: str):
    """Serve a file from the output directory."""
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route('/view/<path:filename>')
def view_file(filename: str):
    """Render an HTML representation of a generated workbook."""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path) or not filename.lower().endswith('.xlsx'):
        return abort(404)
    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as exc:
        logging.error(f"Error reading '{file_path}': {exc}")
        return f"Error reading workbook: {exc}", 500

    sheet_tables: List[Dict[str, Any]] = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_data = build_sheet_structure(ws)
        sheet_tables.append(
            {
                "name": sheet_name,
                "stations": sheet_data["stations"],
                "total_entries": sheet_data["total_entries"],
                "excel_sections": sheet_data["excel_sections"],
            }
        )

    assets = asset_urls()
    return render_template_string(
        """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Viewing {{ filename }}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap">
    <link rel="stylesheet" href="{{ css_black }}">
    <link rel="stylesheet" href="{{ css_custom }}">
    <link rel="stylesheet" href="{{ css_icons }}">
    <style>{{ base_css }}</style>
</head>
<body class="viewer-shell">
<div class="viewer-content">
    <div class="viewer-actions">
        <a class="file-action" href="{{ url_for('index') }}">&larr; Back to dashboard</a>
        <a class="btn btn-primary btn-gradient" href="{{ url_for('download_file', filename=filename) }}">Download workbook</a>
    </div>
    <section class="app-card viewer-hero">
        <div>
            <p class="eyebrow">Online preview</p>
            <h1>{{ filename }}</h1>
            <p class="muted">Review each shift tab without leaving the browser. Use the shift toggles to jump between dayparts or print the active view.</p>
        </div>
        <div class="viewer-meta">
            <div class="stat-chip">
                <span>Total sheets</span>
                <strong>{{ sheets|length }}</strong>
            </div>
            <div class="stat-chip">
                <span>Served from</span>
                <strong>manning_sheets</strong>
            </div>
        </div>
    </section>
    <div class="viewer-tabs">
        {% for table in sheets %}
            <button class="shift-btn{% if loop.first %} active{% endif %}" type="button" data-shift-target="shift-{{ loop.index0 }}">
                <span>{{ table.name }}</span>
                <small>{{ table.total_entries }} assignment{{ "s" if table.total_entries != 1 else "" }}</small>
            </button>
        {% endfor %}
    </div>
    <div class="shift-panels">
    {% for table in sheets %}
        <section class="app-card sheet-card shift-panel{% if loop.first %} active{% endif %}" id="shift-{{ loop.index0 }}">
            <div class="sheet-header">
                <div>
                    <p class="eyebrow">Shift tab</p>
                    <h2>{{ table.name }}</h2>
                </div>
                <div class="panel-actions">
                    <span class="sheet-chip">{{ table.total_entries }} assignment{{ "s" if table.total_entries != 1 else "" }}</span>
                    <button type="button" class="btn btn-sm btn-ghost" data-table-target="table-{{ loop.index0 }}">Show table view</button>
                    <button type="button" class="btn btn-sm btn-gradient" data-print-target="shift-{{ loop.index0 }}">Print Excel View</button>
                </div>
            </div>
            <div class="excel-view">
                <p class="eyebrow">Excel layout</p>
                <table class="excel-table">
                    {% for block in table.excel_sections %}
                        <tr>
                            {% for header in block.headers %}
                                <th>{{ header }}</th>
                            {% endfor %}
                        </tr>
                        <tr>
                            {% for cell in block.cells %}
                                <td>{% if cell %}{{ cell.replace('\\n', '<br>')|safe }}{% else %}&nbsp;{% endif %}</td>
                            {% endfor %}
                        </tr>
                    {% endfor %}
                </table>
            </div>
            <div class="table-scroll table-view" id="table-{{ loop.index0 }}">
                <table class="viewer-table tidy">
                    <thead>
                        <tr>
                            <th>Station</th>
                            <th>Team member</th>
                            <th>Shift</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for station in table.stations %}
                        {% if station.entries %}
                            {% for entry in station.entries %}
                            <tr>
                                {% if loop.first %}
                                <td rowspan="{{ station.entries|length }}">{{ station.station }}</td>
                                {% endif %}
                                <td>{{ entry.name }}</td>
                                <td>{{ entry.time or "—" }}</td>
                            </tr>
                            {% endfor %}
                        {% else %}
                            <tr>
                                <td>{{ station.station }}</td>
                                <td colspan="2" class="muted">No assignments scheduled</td>
                            </tr>
                        {% endif %}
                    {% endfor %}
                    </tbody>
                </table>
            </div>
        </section>
    {% endfor %}
    </div>
</div>
<script src="{{ js_jquery }}"></script>
<script src="{{ js_popper }}"></script>
<script src="{{ js_bootstrap }}"></script>
<script src="{{ js_black }}"></script>
<script>
document.addEventListener('DOMContentLoaded', () => {
    const buttons = Array.from(document.querySelectorAll('[data-shift-target]'));
    const panels = Array.from(document.querySelectorAll('.shift-panel'));
    const activate = (targetId) => {
        panels.forEach(panel => panel.classList.toggle('active', panel.id === targetId));
        buttons.forEach(btn => btn.classList.toggle('active', btn.dataset.shiftTarget === targetId));
    };
    buttons.forEach(btn => {
        btn.addEventListener('click', () => activate(btn.dataset.shiftTarget));
    });
    const printButtons = Array.from(document.querySelectorAll('[data-print-target]'));
    printButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const panel = document.getElementById(btn.dataset.printTarget);
            if (!panel) {
                return;
            }
            const excelView = panel.querySelector('.excel-view');
            if (!excelView) {
                return;
            }
            const printWindow = window.open('', '', 'width=900,height=700');
            if (!printWindow) {
                return;
            }
            printWindow.document.write('<html><head><title>Print Shift</title>');
            printWindow.document.write('<style>@page{size:landscape;margin:10mm;} body{font-family:Poppins,Segoe UI,sans-serif;padding:18px;font-size:12px;} table{width:100%;border-collapse:collapse;table-layout:fixed;} th,td{border:1px solid #000;padding:6px;vertical-align:top;word-break:break-word;} th{background:#f0f0f0;text-transform:uppercase;font-size:0.75rem;} </style>');
            printWindow.document.write('</head><body>');
            printWindow.document.write(`<h2>${panel.querySelector('h2')?.textContent ?? ''}</h2>`);
            printWindow.document.write(excelView.innerHTML);
            printWindow.document.write('</body></html>');
            printWindow.document.close();
            printWindow.focus();
            printWindow.print();
            printWindow.close();
        });
    });
    const tableButtons = Array.from(document.querySelectorAll('[data-table-target]'));
    tableButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const tableView = document.getElementById(btn.dataset.tableTarget);
            if (!tableView) {
                return;
            }
            const isVisible = tableView.classList.toggle('visible');
            btn.textContent = isVisible ? 'Hide table view' : 'Show table view';
        });
    });
    if (buttons.length) {
        activate(buttons[0].dataset.shiftTarget);
    }
});
</script>
</body>
</html>
        """,
        filename=filename,
        sheets=sheet_tables,
        base_css=BASE_CSS,
        **assets,
    )


@app.route('/view_log')
def view_log() -> str:
    """Display the log file in the browser."""
    try:
        with open(log_filename, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as exc:
        return f"Error reading log file: {exc}", 500

    assets = asset_urls()
    return render_template_string(
        """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Processing Log</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap">
    <link rel="stylesheet" href="{{ css_black }}">
    <link rel="stylesheet" href="{{ css_custom }}">
    <link rel="stylesheet" href="{{ css_icons }}">
    <style>{{ base_css }}</style>
</head>
<body class="log-shell">
<div class="viewer-content">
    <div class="viewer-actions">
        <a class="file-action" href="{{ url_for('index') }}">&larr; Back to dashboard</a>
    </div>
    <section class="app-card viewer-hero">
        <div>
            <p class="eyebrow">Diagnostics</p>
            <h1>Processing Log</h1>
            <p class="muted">Trace every upload, validation, and workbook that the generator touched.</p>
        </div>
        <div class="viewer-meta">
            <div class="stat-chip">
                <span>Log file</span>
                <strong>{{ log_name }}</strong>
            </div>
        </div>
    </section>
    <section class="app-card">
        <pre class="log-terminal">{{ content }}</pre>
    </section>
</div>
<script src="{{ js_jquery }}"></script>
<script src="{{ js_popper }}"></script>
<script src="{{ js_bootstrap }}"></script>
<script src="{{ js_black }}"></script>
</body>
</html>
        """,
        content=content,
        base_css=BASE_CSS,
        log_name=os.path.basename(log_filename),
        **assets,
    )


def open_browser():
    """Open the default web browser to the main page."""
    try:
        webbrowser.open("http://127.0.0.1:5000", new=2)
    except Exception:
        # Ignore failures to open browser
        pass


if __name__ == '__main__':
    # Start browser in a separate thread after server launch
    threading.Timer(1.0, open_browser).start()
    logging.info("Starting Manning Chart web server on http://127.0.0.1:5000 ...")
    app.run(host='127.0.0.1', port=5000, debug=False)
