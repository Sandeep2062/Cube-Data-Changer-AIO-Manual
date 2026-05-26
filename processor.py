"""
Cube Data Processor Module
Reads generated data and populates an office template workbook.

Based on: https://github.com/Sandeep2062/Cube-Data-Processor
"""

import os
import shutil
import openpyxl
from openpyxl.utils import column_index_from_string

from generator import generate_rows, grade_display_name, MORTAR_TYPES, ALL_TYPES


# ── Default cell mapping ────────────────────────────────────────────────────

DEFAULT_CELL_MAP = {
    "grade_cell": "B12",
    "casting_date_cell": "C17",
    "date_7d_cell": "C18",
    "date_28d_cell": "F18",
    "weight_row": 25,
    "weight_start_col": "C",
    "weight_count": 6,
    "strength_7d_row": 27,
    "strength_7d_start_col": "C",
    "strength_7d_count": 3,
    "strength_28d_row": 27,
    "strength_28d_start_col": "F",
    "strength_28d_count": 3,
}


def _col_num(col_letter):
    """Convert column letter(s) to 1-based column number."""
    return column_index_from_string(col_letter.strip().upper())


def _get_cell_map(cell_map):
    """Return cell_map merged with defaults."""
    if cell_map is None:
        return dict(DEFAULT_CELL_MAP)
    merged = dict(DEFAULT_CELL_MAP)
    merged.update(cell_map)
    # Normalize user-provided references.
    merged["grade_cell"] = str(merged["grade_cell"]).strip().upper()
    merged["casting_date_cell"] = str(merged["casting_date_cell"]).strip().upper()
    merged["date_7d_cell"] = str(merged["date_7d_cell"]).strip().upper()
    merged["date_28d_cell"] = str(merged["date_28d_cell"]).strip().upper()
    merged["weight_start_col"] = str(merged["weight_start_col"]).strip().upper()
    merged["strength_7d_start_col"] = str(merged["strength_7d_start_col"]).strip().upper()
    merged["strength_28d_start_col"] = str(merged["strength_28d_start_col"]).strip().upper()

    for key in (
        "weight_row", "weight_count",
        "strength_7d_row", "strength_7d_count",
        "strength_28d_row", "strength_28d_count",
    ):
        merged[key] = max(1, int(merged[key]))
    return merged


# ── Helpers ─────────────────────────────────────────────────────────────────

def _normalise_grade_name(raw):
    """Normalise a grade string for matching (strip spaces, uppercase)."""
    return raw.replace(" ", "").upper()


def _extract_grade_from_filename(filepath):
    """Extract grade name from a grade-file filename (legacy support)."""
    name = os.path.basename(filepath).split(".")[0].upper()
    if "MORTAR" in name and "_" in name:
        parts = name.split("_")
        if len(parts) >= 3:
            return f"{parts[-2]}:{parts[-1]}"
    return name.replace("_", "").replace("-", "").strip()


def _load_workbook(filepath):
    """Open a workbook with safe defaults."""
    try:
        return openpyxl.load_workbook(filepath, keep_vba=False, data_only=False, keep_links=False)
    except Exception:
        return openpyxl.load_workbook(filepath)


def _find_sheets_for_grade(office_wb, grade_name, log, cell_map=None):
    """Return list of sheet names whose grade cell matches *grade_name*."""
    cm = _get_cell_map(cell_map)
    grade_cell = cm["grade_cell"]
    target = _normalise_grade_name(grade_name)
    matches = []
    for sheet_name in office_wb.sheetnames:
        ws = office_wb[sheet_name]
        val = ws[grade_cell].value
        if val and _normalise_grade_name(str(val)) == target:
            matches.append(sheet_name)
    return matches


def _grade_from_template_cell(value):
    """Resolve grade/type from template cell B12. Returns None if unsupported."""
    if value is None:
        return None

    raw = str(value).strip().upper()
    if not raw:
        return None

    normalized = raw.replace(" ", "").replace("_", "").replace("-", "")

    # Concrete grades (M10-M45)
    concrete = {g for g in ALL_TYPES if g.startswith("M")}
    if normalized in concrete:
        return normalized

    # Mortar variants commonly seen in templates
    mortar_14_tokens = {"1:4", "1/4", "14", "MORTAR1:4", "MORTAR1/4", "MORTAR14"}
    mortar_16_tokens = {"1:6", "1/6", "16", "MORTAR1:6", "MORTAR1/6", "MORTAR16"}

    if normalized in mortar_14_tokens:
        return "1:4"
    if normalized in mortar_16_tokens:
        return "1:6"

    # Additional fuzzy support
    if "MORTAR" in normalized and ("1:4" in raw or "1/4" in raw or normalized.endswith("14")):
        return "1:4"
    if "MORTAR" in normalized and ("1:6" in raw or "1/6" in raw or normalized.endswith("16")):
        return "1:6"

    return None


# ── Calendar logic ──────────────────────────────────────────────────────────

def load_calendar_data(calendar_file, log):
    """Load calendar Excel → dict[casting_date_str] → {7_days, 28_days}."""
    if not calendar_file or not os.path.exists(calendar_file):
        log("⚠ No calendar file selected")
        return None

    try:
        wb = _load_workbook(calendar_file)
        ws = wb.active
        cal = {}
        row = 2
        while True:
            casting = ws.cell(row=row, column=1).value
            if not casting:
                break
            d7 = ws.cell(row=row, column=2).value
            d28 = ws.cell(row=row, column=3).value
            key = str(casting).strip()
            cal[key] = {
                "7_days": str(d7).strip() if d7 else "",
                "28_days": str(d28).strip() if d28 else "",
            }
            row += 1
        wb.close()
        log(f"✓ Calendar loaded: {len(cal)} dates")
        return cal
    except Exception as e:
        log(f"✖ Calendar error: {e}")
        return None


# ── Date processing ─────────────────────────────────────────────────────────

def apply_dates(office_wb, calendar_data, log, cell_map=None):
    """Write 7-day/28-day dates into every sheet based on casting date cell."""
    cm = _get_cell_map(cell_map)
    casting_cell = cm["casting_date_cell"]
    date_7d_cell = cm["date_7d_cell"]
    date_28d_cell = cm["date_28d_cell"]
    updated = 0
    for sheet_name in office_wb.sheetnames:
        ws = office_wb[sheet_name]
        casting = ws[casting_cell].value
        if not casting:
            continue
        key = str(casting).strip()
        if key in calendar_data:
            d7 = calendar_data[key]["7_days"]
            d28 = calendar_data[key]["28_days"]
            if d7:
                ws[date_7d_cell] = d7
            if d28:
                ws[date_28d_cell] = d28
            updated += 1
            log(f"  ✓ {sheet_name}: {key} → 7d:{d7}, 28d:{d28}")
        else:
            log(f"  ⚠ Date not in calendar: {key} ({sheet_name})")
    return updated


# ── Grade processing (in-memory generation) ─────────────────────────────────

def apply_generated_grades(office_wb, selected_grades, num_rows, log, progress_cb=None, cell_map=None):
    """
    For each selected grade, generate data in-memory and write directly
    into matching sheets of the office workbook.

    Parameters
    ----------
    office_wb : openpyxl.Workbook
    selected_grades : list[str]      e.g. ["M20", "M25", "1:4"]
    num_rows : int                   rows to generate per grade (should >= sheets)
    log : callable
    progress_cb : callable(float)    optional 0-1 progress callback
    cell_map : dict                  optional cell mapping overrides

    Returns total number of sheets populated.
    """
    cm = _get_cell_map(cell_map)
    w_row = int(cm["weight_row"])
    w_col = _col_num(cm["weight_start_col"])
    w_count = int(cm["weight_count"])
    s7_row = int(cm["strength_7d_row"])
    s7_col = _col_num(cm["strength_7d_start_col"])
    s7_count = int(cm["strength_7d_count"])
    s28_row = int(cm["strength_28d_row"])
    s28_col = _col_num(cm["strength_28d_start_col"])
    s28_count = int(cm["strength_28d_count"])

    total = 0
    grade_count = len(selected_grades)

    for gi, grade in enumerate(selected_grades):
        display = grade_display_name(grade)
        sheets = _find_sheets_for_grade(office_wb, grade, log, cell_map)
        log(f"\n  Grade: {display}  →  {len(sheets)} matching sheets")

        if not sheets:
            log(f"  ⚠ No sheets with grade cell = '{grade}'")
            continue

        rows_needed = len(sheets)
        gen = generate_rows(grade, rows_needed)

        for si, (weights, s7d, s28d) in enumerate(gen):
            if si >= len(sheets):
                break
            ws = office_wb[sheets[si]]

            # Weights
            for i, v in enumerate(weights[:w_count]):
                ws.cell(row=w_row, column=w_col + i, value=v)

            # 7-day strengths
            for i, v in enumerate(s7d[:s7_count]):
                ws.cell(row=s7_row, column=s7_col + i, value=v)

            # 28-day strengths
            for i, v in enumerate(s28d[:s28_count]):
                ws.cell(row=s28_row, column=s28_col + i, value=v)

            total += 1
            log(f"    ✓ {sheets[si]} filled")

        if progress_cb:
            progress_cb((gi + 1) / grade_count * 0.8)

    return total


def apply_generated_grades_from_template(office_wb, log, progress_cb=None, cell_map=None):
    """
    Auto mode: read each sheet's grade cell, detect grade/type, group sheets by grade,
    generate rows, and write directly into those sheets.
    """
    from collections import defaultdict
    
    cm = _get_cell_map(cell_map)
    grade_cell = cm["grade_cell"]
    w_row = int(cm["weight_row"])
    w_col = _col_num(cm["weight_start_col"])
    w_count = int(cm["weight_count"])
    s7_row = int(cm["strength_7d_row"])
    s7_col = _col_num(cm["strength_7d_start_col"])
    s7_count = int(cm["strength_7d_count"])
    s28_row = int(cm["strength_28d_row"])
    s28_col = _col_num(cm["strength_28d_start_col"])
    s28_count = int(cm["strength_28d_count"])

    total = 0

    supported_sheets = []
    for sheet_name in office_wb.sheetnames:
        ws = office_wb[sheet_name]
        grade = _grade_from_template_cell(ws[grade_cell].value)
        if grade:
            supported_sheets.append((sheet_name, grade))

    total_supported = len(supported_sheets)
    log(f"  Supported sheets detected from {grade_cell}: {total_supported}")

    if total_supported == 0:
        log(f"  ⚠ No supported grades/types found in {grade_cell} cells")
        return 0

    grade_to_sheets = defaultdict(list)
    for sheet_name, grade in supported_sheets:
        grade_to_sheets[grade].append(sheet_name)

    processed_count = 0
    for grade, sheets in grade_to_sheets.items():
        gen = generate_rows(grade, len(sheets))
        for sheet_name, (weights, s7d, s28d) in zip(sheets, gen):
            ws = office_wb[sheet_name]

            for idx, value in enumerate(weights[:w_count]):
                ws.cell(row=w_row, column=w_col + idx, value=value)
            for idx, value in enumerate(s7d[:s7_count]):
                ws.cell(row=s7_row, column=s7_col + idx, value=value)
            for idx, value in enumerate(s28d[:s28_count]):
                ws.cell(row=s28_row, column=s28_col + idx, value=value)

            total += 1
            processed_count += 1
            log(f"    ✓ {sheet_name} filled ({grade_display_name(grade)})")

            if progress_cb:
                progress_cb(processed_count / total_supported * 0.8)

    return total


# ── Grade processing (from existing Excel files – legacy) ──────────────────

def apply_grade_files(office_wb, grade_files, log, progress_cb=None, cell_map=None):
    """Read existing grade Excel files and populate office template (legacy mode)."""
    cm = _get_cell_map(cell_map)
    w_row = int(cm["weight_row"])
    w_col = _col_num(cm["weight_start_col"])
    w_count = int(cm["weight_count"])
    s7_row = int(cm["strength_7d_row"])
    s7_col = _col_num(cm["strength_7d_start_col"])
    s28_row = int(cm["strength_28d_row"])
    s28_col = _col_num(cm["strength_28d_start_col"])
    s7_count = int(cm["strength_7d_count"])
    s28_count = int(cm["strength_28d_count"])

    total = 0
    file_count = len(grade_files)

    for fi, grade_file in enumerate(grade_files):
        grade_wb = _load_workbook(grade_file)
        grade_ws = grade_wb.active
        grade_name = _extract_grade_from_filename(grade_file)

        log(f"\n  File: {os.path.basename(grade_file)}  (grade: {grade_name})")

        # Find last data row
        row = 2
        while grade_ws.cell(row=row, column=2).value not in (None, ""):
            row += 1
        last_row = row - 1
        log(f"  Data rows: {last_row - 1}")

        sheets = _find_sheets_for_grade(office_wb, grade_name, log, cell_map)
        log(f"  Matching sheets: {len(sheets)}")

        if not sheets:
            grade_wb.close()
            continue

        si = 0
        for r in range(2, last_row + 1):
            if si >= len(sheets):
                log("  ⚠ More data rows than sheets")
                break
            ws = office_wb[sheets[si]]

            weights = [grade_ws.cell(row=r, column=c).value for c in range(2, 8)]
            strengths = [grade_ws.cell(row=r, column=c).value for c in range(9, 15)]

            for i, v in enumerate(weights[:w_count]):
                ws.cell(row=w_row, column=w_col + i, value=v)
            # Split strengths: first s7_count → 7d row, rest → 28d row
            for i, v in enumerate(strengths[:s7_count]):
                ws.cell(row=s7_row, column=s7_col + i, value=v)
            for i, v in enumerate(strengths[s7_count:s7_count + s28_count]):
                ws.cell(row=s28_row, column=s28_col + i, value=v)

            total += 1
            si += 1

        grade_wb.close()

        if progress_cb:
            progress_cb((fi + 1) / file_count * 0.8)

    return total


# ── Main orchestrator ───────────────────────────────────────────────────────

def process(
    office_file,
    output_folder,
    mode,                    # "generate", "grade_files", "date_only", "generate+date", "grade_files+date"
    log,
    selected_grades=None,    # for generate modes
    num_rows=1000,
    grade_files=None,        # for legacy grade-file modes
    calendar_file=None,
    progress_cb=None,
    cell_map=None,           # manual cell mapping overrides
):
    """
    One-shot processing entry point.

    Returns total count of sheet operations performed.
    """
    log(f"\n{'═' * 60}")
    log(f"  MODE: {mode.upper().replace('_', ' ')}")
    log(f"{'═' * 60}")

    # Prepare output
    base = os.path.splitext(os.path.basename(office_file))[0]
    out_name = f"{base}_Processed.xlsx"
    out_path = os.path.join(output_folder, out_name)
    shutil.copy2(office_file, out_path)
    office_wb = _load_workbook(out_path)

    total = 0

    # Calendar
    calendar_data = None
    if "date" in mode:
        calendar_data = load_calendar_data(calendar_file, log)
        if not calendar_data:
            log("✖ Cannot proceed without valid calendar file")
            office_wb.close()
            return 0

    # Grade generation (AIO)
    if "generate" in mode:
        log("\n── GENERATING & APPLYING GRADE DATA ──")
        if selected_grades:
            total += apply_generated_grades(office_wb, selected_grades, num_rows, log, progress_cb, cell_map)
        else:
            cm = _get_cell_map(cell_map)
            log(f"  Auto mode: detecting grade/type from each sheet {cm['grade_cell']}")
            total += apply_generated_grades_from_template(office_wb, log, progress_cb, cell_map)

    # Grade files (legacy)
    if "grade_files" in mode and grade_files:
        log("\n── APPLYING GRADE FILES ──")
        total += apply_grade_files(office_wb, grade_files, log, progress_cb, cell_map)

    # Dates
    if calendar_data:
        log("\n── APPLYING DATES ──")
        updated = apply_dates(office_wb, calendar_data, log, cell_map)
        log(f"  Sheets updated with dates: {updated}")

    # Save
    office_wb.save(out_path)
    office_wb.close()

    log(f"\n{'═' * 60}")
    log(f"  ✓ SAVED → {out_path}")
    log(f"{'═' * 60}")

    if progress_cb:
        progress_cb(1.0)

    return total
