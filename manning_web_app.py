"""
Manning Chart Web Application
============================

This script provides a simple web interface for generating Manning Chart
workbooks from staff scheduling spreadsheets. It supports multiple locations
(e.g., "Ikes" and "Southside") with configurable station mappings.

"""

import os
import re
import sys
import threading
import logging
import webbrowser
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

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
    from openpyxl.utils import get_column_letter
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

# Locations configuration
LOCATIONS = {
    "ikes": {"name": "Ikes", "mapping_needed": False},
    "southside": {"name": "Southside", "mapping_needed": True},
}


from mappings import SOUTHSIDE_MAPPING, IKES_MAPPING, SOUTHSIDE_KEYWORDS, IKES_KEYWORDS

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


# Additional mappings not in the Excel file
FALLBACK_MAPPINGS = {
    "am pasta": "LITTLE ITALY",
}

# Roles to strictly ignore (not count as missing assignments)
IGNORED_ROLES = {
    "scheduled elsewhere",
}


def get_category(role: str, location: str) -> Optional[str]:
    """Map a job role to a Manning Chart category (station) based on location.
    
    1. Exact match (case-insensitive) against dictionary.
    2. Fallback exact matches.
    3. Fuzzy keyword match.
    """
    if not role:
        return None
    role_clean = role.strip().replace("\n", "").strip()
    role_lower = role_clean.lower()

    if location == "southside":
        # 1. Exact lookup
        if role_lower in SOUTHSIDE_MAPPING:
            return SOUTHSIDE_MAPPING[role_lower]
        if role_lower in FALLBACK_MAPPINGS:
            return FALLBACK_MAPPINGS[role_lower]
            
        # 2. Fuzzy/Keyword lookup
        for keyword, station in SOUTHSIDE_KEYWORDS:
            if keyword in role_lower:
                return station
                
        return None

    elif location == "ikes":
        # 1. Exact lookup
        if role_lower in IKES_MAPPING:
            return IKES_MAPPING[role_lower]
            
        # 2. Fuzzy/Keyword lookup
        for keyword, station in IKES_KEYWORDS:
            if keyword in role_lower:
                return station

        return None
    
    return None


def get_stations_layout(location: str) -> List[List[str]]:
    """Return the grid layout of stations for the Manning Sheet."""
    if location == "southside":
        # 5-column layout for Southside
        known_stations = sorted(list(set(SOUTHSIDE_MAPPING.values())))
        # Filter out None or empty
        known_stations = [s for s in known_stations if s]
        
        # Create rows of 5
        rows = []
        chunk_size = 5
        for i in range(0, len(known_stations), chunk_size):
            rows.append(known_stations[i:i + chunk_size])
        
        if not rows:
            rows = [['NO STATIONS MAPPED']]
            
        return rows

    else:
        # Ikes Layout - Dynamic
        known_stations = sorted(list(set(IKES_MAPPING.values())))
        known_stations = [s for s in known_stations if s]

        # Create rows of 5
        rows = []
        chunk_size = 5
        for i in range(0, len(known_stations), chunk_size):
            rows.append(known_stations[i:i + chunk_size])

        if not rows:
             rows = [['NO STATIONS MAPPED']]

        return rows


def process_schedule_file(file_path: str, output_dir: str, location: str = "ikes") -> List[str]:
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

    # Shift definitions based on location
    if location == "southside":
        shifts = [
            {'name': '5.30am-3.00pm', 'meal_periods': 'B L', 'lower': 5.5, 'upper': 15.0},
            {'name': '3.00pm-11.30pm', 'meal_periods': 'D', 'lower': 15.0, 'upper': 23.5},
        ]
    else:
        # Ikes default shifts
        shifts = [
            {'name': '6am-2pm', 'meal_periods': 'B BR', 'lower': 6.0, 'upper': 14.0},
            {'name': '2pm-11pm', 'meal_periods': 'D', 'lower': 14.0, 'upper': 22.0},
            {'name': '10pm-6am', 'meal_periods': 'OVNT', 'lower': 22.0, 'upper': 24.0}, 
        ]

    # Chart row layout
    row_groups = get_stations_layout(location)
    location_name = LOCATIONS.get(location, {}).get("name", location.title())

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
        
        # Metrics for verification
        total_assignments_found = 0
        mapped_assignments_found = 0
        unmapped_roles_list = set()

        for row in ws.iter_rows(min_row=3, values_only=True):
            role = row[0]
            cell_val = row[col_idx]
            if not cell_val:
                continue
            
            category = get_category(str(role), location)
            
            # Count potential assignments in this cell
            cell_str = str(cell_val).strip()
            assignments = [s for s in re.split(r"\n{2,}", cell_str) if s.strip()]
            
            # Temporary list to hold valid assignments found in this cell
            valid_assignments_in_cell = []

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
                
                # Check if valid time parse
                start_time = parse_time(m.group(1))
                if start_time is None:
                    continue
                
                # Check if role is ignored
                if str(role).strip().replace("\n", "").strip().lower() in IGNORED_ROLES:
                    continue

                total_assignments_found += 1
                valid_assignments_in_cell.append((name, m, start_time))

            if not category:
                # If role is not mapped, all assignments in this cell are unmapped
                if valid_assignments_in_cell:
                    unmapped_roles_list.add(str(role))
                continue
            
            # Ensure category exists in layout
            found_layout = False
            for grp in row_groups:
                if category in grp:
                    found_layout = True
                    break
            
            if not found_layout:
                if valid_assignments_in_cell:
                    logging.warning(f"Role '{role}' mapped to '{category}' which is not in the layout.")
                continue

            # Process valid assignments
            for name, m, start_time in valid_assignments_in_cell:
                # Assign to shift based on start time
                shift_index = -1
                
                if location == 'southside':
                    # Simple range check
                    for idx, s in enumerate(shifts):
                        if s['lower'] <= start_time < s['upper']:
                            shift_index = idx
                            break
                        # Handle potential edge case where 11:30pm might be exactly 23.5
                else:
                    # Ikes logic (legacy behavior preservation)
                    if 6 <= start_time < 14:
                        shift_index = 0
                    elif 14 <= start_time < 22:
                        shift_index = 1
                    else:
                        shift_index = 2

                if shift_index != -1:
                    # check if category in that shift's dict (it should be initialized for all)
                    if category in shift_data[shift_index]:
                        entry = f"{name}\n{m.group(1)} - {m.group(2)}"
                        shift_data[shift_index][category].append(entry)
                        mapped_assignments_found += 1
        
        logging.info(f"Verification for {date_label}: Found {total_assignments_found} assignments. Mapped {mapped_assignments_found}.")
        if unmapped_roles_list:
            logging.warning(f"Unmapped roles with assignments: {list(unmapped_roles_list)}")
        if total_assignments_found != mapped_assignments_found:
             logging.warning(f"Mismatch in assignment counts! Missing {total_assignments_found - mapped_assignments_found} assignments.")

        # Create output workbook
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)
        for idx_shift, shift_info in enumerate(shifts):
            sheet = out_wb.create_sheet(shift_info['name'])
            # Calculate max columns dynamically
            max_cols = max(len(grp) for grp in row_groups) if row_groups else 1
            if max_cols < 3: max_cols = 3 
            
            from openpyxl.utils import get_column_letter
            end_col_letter = get_column_letter(max_cols)

            sheet.merge_cells(f'A1:{end_col_letter}1')
            sheet['A1'] = f'MANNING CHART - {location_name.upper()}'
            sheet['A1'].font = Font(size=14, bold=True)
            sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
            
            sheet.merge_cells(f'A2:{end_col_letter}2')
            sheet['A2'] = f"Date: {date_label}{weekday_label}    Meal Periods: {shift_info['meal_periods']}    MOD:"
            sheet['A2'].alignment = Alignment(horizontal='left', vertical='center')
            
            for c_idx in range(1, max_cols + 1):
                sheet.column_dimensions[get_column_letter(c_idx)].width = 30
            
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
                
        out_filename = f"{date_part}_{location_name}_Manning_sheet_{generation_stamp}.xlsx"
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
.location-toggle {
    margin-bottom: 30px;
    display: flex;
    justify-content: center;
    gap: 0;
}
.location-toggle .btn {
    min-width: 160px;
    border-radius: 0;
    border: 1px solid rgba(255,255,255,0.1);
}
.location-toggle .btn:first-child {
    border-top-left-radius: 30px;
    border-bottom-left-radius: 30px;
}
.location-toggle .btn:last-child {
    border-top-right-radius: 30px;
    border-bottom-right-radius: 30px;
}
.location-toggle .btn-secondary {
    background: transparent;
    color: rgba(255,255,255,0.7);
}
.location-toggle .btn-primary {
    background: #e14eca;
    background-image: linear-gradient(to bottom left, #e14eca, #ba54f5, #e14eca);
    background-size: 210% 210%;
    background-position: top right;
    border-color: transparent;
    box-shadow: 0px 0px 20px 0px rgba(186, 84, 245, 0.5);
}
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
        
        if header_row and data_row is not None:
             # Find how many columns have content in header
            col_count = len(header_row)
            
            headers: List[str] = []
            cells: List[str] = []
            
            has_content = False
            for col_idx in range(col_count):
                header_val = header_row[col_idx]
                station_name = str(header_val).strip() if header_val else ""
                
                # Simple heuristic: stop if several empty headers in a row or use all?
                # We'll rely on the fact that these are generated sheets with borders.
                
                headers.append(station_name)
                
                cell_val = data_row[col_idx] if col_idx < len(data_row) else None
                assignment_text = str(cell_val) if cell_val else ""
                cells.append(assignment_text)
                
                if station_name:
                    has_content = True
                    parsed = parse_cell_assignments(cell_val)
                    stations.append({"station": station_name, "entries": parsed})

            if has_content:
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
    location = request.args.get("location", "ikes") # Default to Ikes
    
    # Validate location
    if location not in LOCATIONS:
        location = "ikes"
    
    show_history = view_mode == "history"
    all_files = list_output_files() if show_history else CURRENT_OUTPUTS.copy()
    
    # Filter files by location
    if location == "southside":
        files = [f for f in all_files if "_Southside_" in f]
    else:
        # Ikes files might explicitly say Ikes or might be legacy (no location name?)
        # Current logic adds location_name to filename: "{date_part}_{location_name}_Manning_sheet..."
        # So Ikes files should have "_Ikes_".
        files = [f for f in all_files if "_Ikes_" in f or ("_Southside_" not in f)]
    
    total_generated = len(CURRENT_OUTPUTS)
    latest_file = CURRENT_OUTPUTS[-1] if CURRENT_OUTPUTS else None
    assets = asset_urls()
    flashes = get_flashed_messages(with_categories=True)
    
    location_name = LOCATIONS[location]["name"]
    title = f"Manning Sheets {location_name}"

    return render_template_string(
        """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>{{ title }}</title>
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
        
        <div class="location-toggle">
            <div class="btn-group" role="group" aria-label="Location Toggle">
                <a href="{{ url_for('index', location='ikes', view=view_mode) }}" class="btn btn-{{ 'primary' if location == 'ikes' else 'secondary' }}">Manning Sheets Ikes</a>
                <a href="{{ url_for('index', location='southside', view=view_mode) }}" class="btn btn-{{ 'primary' if location == 'southside' else 'secondary' }}">Manning Sheets Southside</a>
            </div>
        </div>

        <div class="hero-row">
            <div class="hero-cell col-3 hero-desc">
                <p class="eyebrow">Automated coverage</p>
                <h2>{{ title }}</h2>
                <p class="muted">Upload the latest <strong>.xlsx</strong> schedule for <strong>{{ location_name }}</strong>. We keep the raw file in <code>input_my_staff_schedule</code> and save processed shifts inside <code>manning_sheets</code>.</p>
            </div>
            <div class="hero-cell col-6 hero-upload">
                <p class="eyebrow text-center">Upload schedule</p>
                <h2 class="text-center">Generate new Manning sheets</h2>
                <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data" class="upload-form centered">
                    <input type="hidden" name="location" value="{{ location }}">
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
                    <a class="btn btn-sm btn-outline" href="{{ url_for('index', view='history', location=location) }}">History</a>
                    <a class="btn btn-sm btn-gradient" href="{{ url_for('index', view='current', location=location) }}">Current</a>
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
        title=title,
        location=location,
        location_name=location_name,
        view_mode=view_mode,
        **assets,
    )


@app.route('/upload', methods=['POST'])
def upload() -> str:
    """Handle file upload, save to input directory, process it, and redirect."""
    global CURRENT_OUTPUTS
    
    location = request.form.get("location", "ikes")
    
    if 'file' not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for('index', location=location))
    file = request.files['file']
    if file.filename == '':
        flash("Please choose a file before uploading.", "error")
        return redirect(url_for('index', location=location))
    if not file.filename.lower().endswith('.xlsx'):
        flash("Only .xlsx files are supported. Upload the MyStaff shift schedule exported in Task Wise view (.xlsx).", "error")
        return redirect(url_for('index', location=location))
    # Save the uploaded file with a timestamp to avoid collisions
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_filename = f"{timestamp}_{os.path.basename(file.filename)}"
    input_path = os.path.join(INPUT_DIR, safe_filename)
    file.save(input_path)
    logging.info(f"Uploaded schedule saved to '{input_path}'.")
    # Process the uploaded schedule
    outputs = process_schedule_file(input_path, OUTPUT_DIR, location=location)
    if outputs:
        logging.info(f"Generated {len(outputs)} output file(s) from '{safe_filename}'.")
        CURRENT_OUTPUTS = outputs
        flash(f"Successfully generated {len(outputs)} chart(s).", "success")
    else:
        logging.warning(f"No output files generated from '{safe_filename}'.")
        flash("Unable to process that file. Upload the MyStaff shift schedule exported in Task Wise view (.xlsx) and try again.", "error")
    # Redirect to index to show new files
    return redirect(url_for('index', location=location))


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
                <div style="overflow-x: auto;">
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
        with open(log_filename, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception as exc:
        content = f"Error reading log: {exc}"
    return render_template_string(
        """
<!doctype html>
<html>
<head>
    <title>System Log</title>
    <style>body{background:#1e1e2f;color:#e1e4e8;font-family:monospace;padding:20px;white-space:pre-wrap;}</style>
</head>
<body>{{ content }}</body>
</html>
        """,
        content=content,
    )


if __name__ == '__main__':
    open_browser = True
    if len(sys.argv) > 1 and sys.argv[1] == '--no-browser':
        open_browser = False

    # Start the server
    port = int(os.environ.get("PORT", 5000))
    if open_browser:
        threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    
    app.run(host='0.0.0.0', port=port)
