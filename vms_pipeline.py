#!/usr/bin/env python3
"""
Linways Attendance -> VMS Report — ALL-IN-ONE SCRIPT
=======================================================

Two things this script can do:

  1) FORMAT a raw Linways export you already downloaded into the full VMS
     report (GEN / GEN ALL / per-section sheets / SUMMARY / bracket tables
     / color coding) — works right now, no login needed.

  2) DOWNLOAD the raw report from Linways automatically (Selenium), then
     format it — needs a few CSS selectors filled in first (see the
     SELECTOR CONFIG section below and the instructions at the bottom of
     this file).

USAGE
-----
Just format a file you already have:
    python vms_pipeline.py format raw_attendance.xlsx VMS_Report.xlsx --low 0 --high 75

    Limit to specific subjects only (whitelist):
        --include "Data Structures,Operating Systems"
    Drop specific subjects (blacklist, applied after --include):
        --exclude "Soft Skill,Yoga"

Download from Linways AND format in one go (after filling in selectors):
    python vms_pipeline.py pipeline --username YOU --password PASS --output VMS_Report.xlsx \
        --from-date 01/06/2026 --to-date 31/07/2026

    --from-date/--to-date filter the report AT THE SOURCE (Linways' own date
    range field on the report page) before it's downloaded — they need
    SELECTORS["date_from"] / SELECTORS["date_to"] filled in first, same as
    the other selectors (see instructions at the bottom of this file).

Only download the raw file (no formatting):
    python vms_pipeline.py download --username YOU --password PASS --from-date 01/06/2026 --to-date 31/07/2026

NOTE: the Abstract workbook NEVER includes Soft Skill, regardless of the
Soft Skill choice made for the main VMS report / debar lists.
"""
import argparse
import glob
import os
import re
import time
from collections import Counter

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


# ══════════════════════════════════════════════════════════════════════
# SELECTOR CONFIG — placeholders, fill in with real values (see bottom)
# ══════════════════════════════════════════════════════════════════════
LOGIN_URL = "https://presidencycollege.linways.com/ams/faculty/login"
DOWNLOAD_DIR = os.path.abspath("./downloads")

# These need to be updated from DevTools -> Inspect on the real Linways site.
# Import selenium's By lazily so `format` mode works without selenium installed.
#
# Confirmed via user-supplied outerHTML (Vue.js SPA, note the data-v-* scoped
# attrs — these hashes can change between Linways deployments/updates, so if
# selectors break again after a Linways update, re-copy the outerHTML).
BASE_URL = "https://presidencycollege.linways.com"
REPORT_PATH = "/ams/faculty/attendance/consolidated-course-wise-report?redir=true"

SELECTORS = {
    "username": ("NAME", "username"),                                   # (By kind, value)
    "password": ("NAME", "password"),
    "login_button": ("ID", "loginBtn"),
    "search_button": ("XPATH", "//span[normalize-space(text())='Search']"),
    "export_dropdown_toggle": ("XPATH", "//span[normalize-space(text())='Export']"),
    "download_button": ("XPATH", "//a[contains(@class,'dropdown-item') and normalize-space(text())='Excel']"),

    # PLACEHOLDERS — this is almost certainly the "unlabeled input field
    # (form-control)" mentioned in download_consolidated_report() below.
    # Inspect it on the real report page and update these (it may turn out
    # to be ONE combined daterangepicker field rather than two separate
    # ones — if so, point both entries at the same selector and adjust
    # set_date_range() below to type "from - to" into it in one go).
    "date_from": ("XPATH", "//input[@placeholder='From Date']"),
    "date_to": ("XPATH", "//input[@placeholder='To Date']"),
}


# ══════════════════════════════════════════════════════════════════════
# PART 1 — REPORT FORMATTING (mirrors your VMS_Gold.html logic exactly)
# ══════════════════════════════════════════════════════════════════════
IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA",
          "TABLE TENNIS", "FREESLOT", "FREE SLOT", "ATOM", "DSA", "LIB", "LIBRARY", "MENTORING"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"
LAB_RE = re.compile(r"(LAB|PRACTICAL|WORKSHOP)", re.I)

RED_FILL = PatternFill(fgColor="C0392B", fill_type="solid")
GREEN_FILL = PatternFill(fgColor="C6EFCE", fill_type="solid")
HEADER_FILL = PatternFill(fgColor="2C3E50", fill_type="solid")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=11)
BLACK_BOLD = Font(bold=True, color="000000", size=11)
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
PLAIN_FONT = Font(size=11)
THIN = Side(style="thin", color="4D4D4D")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Subjects containing this keyword (case-insensitive) are dropped when
# Soft Skill is excluded. Defined up here so every builder can use it.
SOFT_SKILL_KEYWORD = "SOFT SKILL"


def is_valid(subject, extra_ignore=None):
    if subject is None:
        return False
    s = str(subject).strip()
    if not s:
        return False
    u = s.upper()
    if any(k in u for k in IGNORE):
        return False
    if extra_ignore and any(k in u for k in extra_ignore):
        return False
    return True


def r2(v):
    return round(float(v), 2)


def load_raw(path, extra_ignore=None):
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    raw = [[c.value for c in row] for row in ws.iter_rows()]

    h_row = 0
    for i in range(min(15, len(raw))):
        if any(x is not None and "ROLL NO" in str(x).upper() for x in raw[i]):
            h_row = i
            break
    hdrs = raw[h_row]

    C = {}
    for h in hdrs:
        if h is None:
            continue
        hs = str(h).strip()
        if re.search(r"Roll No", hs, re.I):
            C["roll"] = h
        if re.search(r"Student Name", hs, re.I):
            C["name"] = h
        if re.fullmatch(r"Batch", hs, re.I):
            C["batch"] = h
        if re.search(r"Semester", hs, re.I):
            C["sem"] = h
        if re.search(r"Course Name|Subject Name", hs, re.I):
            C["subject"] = h
        if hs == ATT_COL_NAME:
            C["attendance"] = h
        if re.search(r"Faculty", hs, re.I):
            C["faculty"] = h

    C.setdefault("roll", hdrs[1])
    C.setdefault("name", hdrs[2])
    C.setdefault("sem", hdrs[5])
    C.setdefault("batch", hdrs[6])
    C.setdefault("subject", hdrs[8])
    C.setdefault("attendance", hdrs[15])
    # Faculty name: column Q (17th column, index 16) in the consolidated
    # report, unless a header explicitly named "Faculty..." was found above.
    C.setdefault("faculty", hdrs[16] if len(hdrs) > 16 else None)

    rows = []
    for r in range(h_row + 1, len(raw)):
        row = raw[r]
        roll_v = row[1] if len(row) > 1 else None
        name_v = row[2] if len(row) > 2 else None
        if (roll_v is None or str(roll_v).strip() == "") and (name_v is None or str(name_v).strip() == ""):
            continue
        obj = {}
        for i, h in enumerate(hdrs):
            if h is None:
                continue
            obj[h] = row[i] if i < len(row) else None
        rows.append(obj)

    G = []
    for r in rows:
        if not is_valid(r.get(C["subject"]), extra_ignore):
            continue
        batch_val = str(r.get(C["batch"]) or "")
        r["_dept"] = batch_val.strip().split(" ")[0].upper() if batch_val.strip() else ""
        G.append(r)
    return G, C


def canon_section(s):
    """Normalize a section string purely for MATCHING purposes — never
    for anything actually displayed or written into a sheet.

    The consolidated report's own Batch column and the Lab Batch List's
    Batch column don't always spell the same section the same way — e.g.
    the report says 'BCA AIML 2026 - A' but the Batch List says
    'BCA AIML 2026 A' (no dash). Compared as plain strings those never
    match, so the lab/faculty/batch cross-reference silently misses every
    section spelled that way — this is exactly what used to make the
    Batch List's batch numbers not apply to the BCA AIML / BCA DS
    sections. canon_section() collapses dashes and extra whitespace and
    uppercases, so both spellings map to the same key."""
    return re.sub(r"[\s\-]+", " ", str(s).strip()).strip().upper()


def series_of(batch):
    """'BCA 2025 A' -> 'BCA 2025'; 'BCA AIML 2026 - A' -> 'BCA AIML 2026'
    — the series (full program/specialization stream + year) a section
    belongs to, independent of the trailing section letter. Takes every
    token up to and including the 4-digit year, so a specialization
    stream like AIML or DS gets its OWN series instead of being folded
    into the plain BCA series for the same year — 'BCA AIML 2026 - A'
    and 'BCA 2026 A' must never be treated as the same GEN/GEN ALL
    group just because they share a year."""
    parts = str(batch).strip().split()
    if not parts:
        return ""
    year_idx = next((i for i, p in enumerate(parts) if re.fullmatch(r"\d{4}", p)), None)
    if year_idx is None:
        # no clean 4-digit year token — fall back to first token + first
        # numeric token (or 'X') so odd/legacy batch strings still
        # produce *something* usable instead of misgrouping silently.
        yr = next((p for p in parts if p.isdigit()), "X")
        return f"{parts[0]} {yr}"
    return " ".join(parts[: year_idx + 1])


def get_series(dept_rows, C):
    batches = {str(r[C["batch"]]) for r in dept_rows}
    return sorted({series_of(b) for b in batches})


def series_df(dept_rows, ser, C):
    """Every row whose section belongs EXACTLY to series `ser` — matched
    via series_of(), not a loose substring test (the old
    'k0 in batch and k1 in batch' check would match "2026" appearing
    anywhere in the string, which is how 'BCA AIML 2026 - A' used to leak
    into the plain 'BCA 2026' series)."""
    return [r for r in dept_rows if series_of(r[C["batch"]]) == ser]


def get_subjects(rows, C, extra_ignore=None):
    return sorted({r[C["subject"]] for r in rows if is_valid(r[C["subject"]], extra_ignore)})


# ══════════════════════════════════════════════════════════════════════
# PART 1B — BCA LAB BATCH LIST CROSS-REFERENCE (labs & internship)
# ══════════════════════════════════════════════════════════════════════
COURSE_CODE_RE = re.compile(r"^\S+\s*-\s*")
BATCH_SUFFIX_RE = re.compile(r"\s+(BATCH|Batch|batch)\s*(\d+)\s*$")


def parse_course_community_name(name):
    """'SGC201.1P - C Programming Lab BATCH 1' -> ('C Programming Lab', 'Batch 1')
    'SGC205.5C - INTERNSHIP WITH MINI PROJECT' -> ('INTERNSHIP WITH MINI PROJECT', None)
    Strips the leading course code and trailing batch number, matching the
    subject text exactly as it appears in the consolidated report's own
    Course Name column."""
    name = str(name).strip()
    name = COURSE_CODE_RE.sub("", name, count=1).strip()
    m = BATCH_SUFFIX_RE.search(name)
    if m:
        subject = BATCH_SUFFIX_RE.sub("", name).strip()
        return subject, f"Batch {m.group(2)}"
    return name, None


def load_batch_list(path):
    """Read every sheet of a BCA/MCA Lab Batch List workbook into a flat
    list of (section, subject, roll, batch_label, faculty) tuples — one
    per student per lab/internship. `section` is sanitized the same way
    write_sheet() sanitizes section names, so it lines up with the debar
    sheet titles and the consolidated report's own Batch column."""
    wb = load_workbook(path, data_only=True)
    entries = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        headers = [c.value for c in ws[1]]

        def idx(name):
            return headers.index(name) if name in headers else -1

        i_roll = idx("Reg No.")
        i_section = idx("Batch")
        i_course = idx("Course Community Name")
        i_fac = idx("Faculty")
        if -1 in (i_roll, i_section, i_course, i_fac):
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[i_roll] is None:
                continue
            roll = str(row[i_roll]).strip()
            section = str(row[i_section]).strip().replace("/", "-") if row[i_section] else ""
            course_raw = row[i_course]
            fac = str(row[i_fac]).strip() if row[i_fac] else ""
            if not (roll and section and course_raw and fac):
                continue
            subject, batch_label = parse_course_community_name(course_raw)
            # Roll 25CG102 (PRUTHVI RAJ H)'s Batch List row for Operating
            # Systems Lab, BCA 2025 Section B, has a bad/merged faculty
            # entry ("Nasrulla Khan K,PARVEZ AHMED SHARIFF") — force
            # Parvez for THIS roll + THIS section + THIS subject only.
            # The student's other two labs are left untouched, using
            # whatever faculty the sheet says.
            if (
                roll.upper() == "25CG102"
                and section.upper() == "BCA 2025 B"
                and subject.strip().upper() == "OPERATING SYSTEMS LAB"
            ):
                fac = "PARVEZ AHMED SHARIFF"
            entries.append((section, subject, roll, batch_label, fac))
    return entries


ROMAN_SEM = {1: "I", 2: "II", 3: "III", 4: "IV", 5: "V", 6: "VI", 7: "VII", 8: "VIII"}


def sem_label(sem_str):
    """'S1' -> 'I', 'S3' -> 'III'."""
    m = re.search(r"\d+", str(sem_str))
    if not m:
        return str(sem_str)
    return ROMAN_SEM.get(int(m.group()), str(m.group()))


ABSTRACT_BUCKETS = [
    (0.00, 29.99), (30.00, 49.99), (50.00, 59.99),
    (60.00, 64.49), (64.50, 69.99), (70.00, 74.99),
]
ABSTRACT_BUCKET_LABELS = ["0.00-29.99", "30.00-49.99", "50.00-59.99",
                           "60.00-64.49", "64.50-69.99", "70.00-74.99"]


def bucket_index(pct):
    """Which of the 6 shortage bands (0-74.99%) `pct` falls into, or None
    if `pct` is >=75% (not a shortage case) or unparseable."""
    try:
        p = round(float(pct), 2)
    except (TypeError, ValueError):
        return None
    for i, (lo, hi) in enumerate(ABSTRACT_BUCKETS):
        if lo - 1e-9 <= p <= hi + 1e-9:
            return i
    return None


def process_grid(rows, subjects, low_t, high_t, show_all, C):
    if not rows:
        return None, None

    pivot = {}
    for r in rows:
        key = str(r[C["roll"]])
        if key not in pivot:
            pivot[key] = {
                C["roll"]: r[C["roll"]],
                C["name"]: r[C["name"]],
                C["batch"]: r[C["batch"]],
                C["sem"]: r[C["sem"]],
            }
        v = r.get(C["attendance"])
        try:
            v = r2(v)
        except (TypeError, ValueError):
            v = None
        pivot[key][r[C["subject"]]] = v

    prows = list(pivot.values())

    lab_re = LAB_RE
    for row in prows:
        th_subs = [s for s in subjects if not lab_re.search(s)]
        th_vals = [row[s] for s in th_subs if row.get(s) is not None]
        row["Theory Avg"] = r2(sum(th_vals) / len(th_vals)) if th_vals else None
        all_vals = [row[s] for s in subjects if row.get(s) is not None]
        row["Final Avg"] = r2(sum(all_vals) / len(all_vals)) if all_vals else None

    if not show_all:
        prows = [
            row for row in prows
            if any(row.get(s) is not None and low_t <= row[s] <= high_t for s in subjects)
        ]
    if not prows:
        return None, None

    counts = {s: 0 for s in subjects}
    for row in prows:
        cnt = 0
        for s in subjects:
            v = row.get(s)
            if v is not None and low_t <= v <= high_t:
                cnt += 1
                counts[s] += 1
        row["Subjects in Range"] = cnt
        if not show_all:
            for s in subjects:
                v = row.get(s)
                if v is None or v < low_t or v > high_t:
                    row[s] = ""

    for i, row in enumerate(prows):
        row["Sl No."] = i + 1

    count_row = {"Sl No.": "", "_isCount": True,
                 C["roll"]: "", C["name"]: "", C["batch"]: "",
                 C["sem"]: f"Count ({low_t}-{high_t}%)"}
    for s in subjects:
        count_row[s] = counts[s]
    count_row["Theory Avg"] = ""
    count_row["Final Avg"] = ""
    count_row["Subjects in Range"] = ""
    prows.append(count_row)

    return prows, counts


def get_bracket(rows, subjects, threshold, C):
    out = []
    for sub in subjects:
        vals = []
        for r in rows:
            if r[C["subject"]] != sub:
                continue
            v = r.get(C["attendance"])
            try:
                vals.append(r2(v))
            except (TypeError, ValueError):
                pass
        row = {"Subject": sub}
        tot = 0
        if threshold > 0:
            n = sum(1 for v in vals if 0 <= v < 50)
            row["0.00-49.99"] = n; tot += n
        if threshold > 50:
            n = sum(1 for v in vals if 50 <= v < 60)
            row["50.00-59.99"] = n; tot += n
        if threshold > 60:
            a = sum(1 for v in vals if 60 <= v < 64.5)
            b = sum(1 for v in vals if 64.5 <= v < 70)
            row["60.00-64.49"] = a
            row["64.50-69.99"] = b
            tot += a + b
        if threshold > 70:
            n = sum(1 for v in vals if 70 <= v < 75)
            row["70.00-74.99"] = n; tot += n
        row["Total"] = tot
        out.append(row)
    return out


def write_sheet(wb, sname, grid_rows, raw_rows, subjects, threshold, C):
    ws = wb.create_sheet(title=sname[:31])

    cols = ["Sl No.", C["roll"], C["name"], C["batch"], C["sem"], *subjects,
            "Theory Avg", "Final Avg", "Subjects in Range"]
    sub_start = 5
    sub_end = 5 + len(subjects) - 1
    th_avg_col = sub_end + 1
    fn_avg_col = sub_end + 2
    sir_col = sub_end + 3

    for ci, c in enumerate(cols):
        cell = ws.cell(row=1, column=ci + 1, value=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left" if ci == 2 else "center",
                                    vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 62.4

    widths = [5, 13, 41, 12, 8] + [20] * len(subjects) + [12, 12, 14]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    for r_idx, row in enumerate(grid_rows, start=2):
        is_count = row.get("_isCount", False)
        for ci, c in enumerate(cols):
            v = row.get(c, "")
            if v is None:
                v = ""
            cell = ws.cell(row=r_idx, column=ci + 1, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="left" if ci == 2 else "center", vertical="center")
            cell.font = PLAIN_FONT

            if is_count:
                if sub_start <= ci <= sub_end:
                    try:
                        n = float(v)
                    except (TypeError, ValueError):
                        n = None
                    if n is not None and n > 0:
                        cell.fill = RED_FILL
                        cell.font = WHITE_BOLD
            else:
                try:
                    n = float(v)
                    has_val = v != "" and v is not None
                except (TypeError, ValueError):
                    n = None
                    has_val = False

                if sub_start <= ci <= sub_end and has_val:
                    if n < 70:
                        cell.fill = RED_FILL; cell.font = WHITE_BOLD
                    elif n < threshold:
                        cell.fill = GREEN_FILL; cell.font = BLACK_BOLD

                if ci in (th_avg_col, fn_avg_col) and has_val:
                    if n < 70:
                        cell.fill = RED_FILL; cell.font = WHITE_BOLD
                    elif n < threshold:
                        cell.fill = GREEN_FILL; cell.font = BLACK_BOLD

                if ci == sir_col and has_val and n > 0:
                    cell.fill = RED_FILL; cell.font = WHITE_BOLD

    bdata = get_bracket(raw_rows, subjects, threshold, C)
    bkt_hdr_r = len(grid_rows) + 3
    if bdata:
        bkeys = list(bdata[0].keys())
        for ci, k in enumerate(bkeys):
            cell = ws.cell(row=bkt_hdr_r, column=ci + 1, value=k)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
        for ri, brow in enumerate(bdata):
            for ci, k in enumerate(bkeys):
                v = brow[k]
                cell = ws.cell(row=bkt_hdr_r + 1 + ri, column=ci + 1, value=v)
                cell.border = BORDER
                cell.alignment = Alignment(horizontal="left" if ci == 0 else "center", vertical="center")
                cell.font = PLAIN_FONT
                if ci > 0 and isinstance(v, (int, float)) and v > 0 and k in ("Total", "70.00-74.99"):
                    cell.fill = RED_FILL
                    cell.font = WHITE_BOLD

    return ws


def write_summary_sheet(wb, summaries):
    ws = wb.create_sheet(title="SUMMARY")
    ws.cell(row=1, column=1, value="Section").fill = HEADER_FILL
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=2, value="Count").fill = HEADER_FILL
    ws.cell(row=1, column=2).font = HEADER_FONT
    for ci in (1, 2):
        ws.cell(row=1, column=ci).border = BORDER
        ws.cell(row=1, column=ci).alignment = Alignment(horizontal="center", vertical="center")
    for i, s in enumerate(summaries, start=2):
        ws.cell(row=i, column=1, value=s["Section"]).border = BORDER
        ws.cell(row=i, column=2, value=s["Count"]).border = BORDER
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10


def prompt_for_range(low, high):
    """Ask for the attendance % range interactively if not given via CLI.

    Lets you choose the limit fresh each time you run the script, instead
    of relying on a hard-coded default. Passing --low/--high on the command
    line still skips the prompt for that one.
    """
    if low is None:
        while True:
            raw = input("Attendance % lower limit (--low) [default 0]: ").strip()
            if raw == "":
                low = 0.0
                break
            try:
                low = float(raw)
                break
            except ValueError:
                print("  Enter a number, e.g. 0")
    if high is None:
        while True:
            raw = input("Attendance % upper limit (--high) [default 75]: ").strip()
            if raw == "":
                high = 75.0
                break
            try:
                high = float(raw)
                break
            except ValueError:
                print("  Enter a number, e.g. 75")
    return low, high


def prompt_for_soft_skill(flag):
    """Ask whether to include Soft Skill in the report if not given via CLI.

    flag is True/False/None: None means "not specified on the command
    line" -> prompt; True/False (from --include-soft-skill/--exclude-soft-skill)
    skips the prompt.
    """
    if flag is not None:
        return flag
    while True:
        raw = input("Include Soft Skill in the report? [y/N]: ").strip().lower()
        if raw in ("", "n", "no"):
            return False
        if raw in ("y", "yes"):
            return True
        print("  Enter y or n")


def apply_subject_filters(rows, C, exclude=None, include=None):
    """Whitelist (include) then blacklist (exclude) rows by subject name.

    include: if non-empty, ONLY these subjects are kept.
    exclude: applied after include, drops any of these subjects that remain.
    Both are matched case-sensitively against the subject column exactly as
    it appears in the raw file (same as the original --exclude behavior).
    """
    include = set(include or [])
    exclude = set(exclude or [])
    out = rows
    if include:
        out = [r for r in out if r[C["subject"]] in include]
    if exclude:
        out = [r for r in out if r[C["subject"]] not in exclude]
    return out


# Hidden sheets used to carry faculty/group info from the VMS Report
# workbook through to the debar-list legend.
FACULTY_MAP_SHEET = "_FacultyMap"          # (section, subject) -> list of (batch_label, faculty) groups
STUDENT_FACULTY_MAP_SHEET = "_FacultyStudentMap"  # per-student, multi-group subjects only

# A GROUP is the real unit that can need its own legend/abstract row: the
# pair (batch_label, faculty). Two different lab batches taught by the
# SAME faculty are two different groups and must stay on separate rows —
# grouping by faculty name alone would silently merge their shortage
# counts together. batch_label is None for theory subjects / subjects the
# Batch List doesn't cover, where faculty name alone is the group.


def group_label(batch_label, faculty):
    """Display string for a group: 'Batch 1 - NAME' when a batch label is
    known, else just the faculty name."""
    return f"{batch_label} - {faculty}" if batch_label else faculty


def group_sort_key(group):
    """Sort (batch_label, faculty) groups by batch NUMBER (1, 2, 3 …),
    not faculty name — so batches always appear in chronological order
    (Batch 1, Batch 2, Batch 3 ...) instead of alphabetically by whoever
    teaches them. Groups with no batch label sort first."""
    batch_label, fac = group
    if batch_label:
        m = re.search(r"\d+", batch_label)
        n = int(m.group()) if m else 0
    else:
        n = 0
    return (n, fac)


def build_faculty_maps(rows, C, batch_entries=None):
    """Two maps, built in one pass over the raw rows plus an optional
    cross-reference against a BCA/MCA Lab Batch List (`batch_entries`,
    from load_batch_list()):

    subject_groups: (section, subject) -> sorted list of every DISTINCT
    (batch_label, faculty) GROUP teaching that subject in that section.
    Deliberately per-section, keeping every distinct group rather than
    picking one, because real timetables aren't uniform — including two
    lab batches taught by the same person, which must stay separate. For
    any (section, subject) that appears in `batch_entries` (i.e. a
    lab/internship the Batch List covers), the Batch List is authoritative
    and fully replaces whatever the consolidated report's own faculty
    column said for that subject — the Batch List is curated specifically
    for lab-batch assignment, the report's faculty column is not.

    student_group: (section, subject, roll) -> (batch_label, faculty),
    kept ONLY for subjects with more than one distinct group in that
    section — i.e. exactly where the legend needs to split into one row
    per group.
    """
    fac_key = C.get("faculty")

    mapping = {}  # (canon section, subject) -> set of (batch_label, faculty) groups
    entries = []  # (canon section, subject, roll, batch_label, faculty)
    if fac_key:
        for r in rows:
            subj = r.get(C["subject"])
            fac = r.get(fac_key)
            batch = r.get(C["batch"])
            roll = r.get(C["roll"])
            if not subj or fac is None or not batch or roll in (None, ""):
                continue
            subj = str(subj).strip()
            fac = str(fac).strip()
            # match the same sanitizing write_sheet() applies to sheet names
            batch = str(batch).strip().replace("/", "-")
            roll = str(roll).strip()
            if not (subj and fac and batch and roll):
                continue
            key_section = canon_section(batch)
            mapping.setdefault((key_section, subj), set()).add((None, fac))
            entries.append((key_section, subj, roll, None, fac))

    # Batch List cross-reference: authoritative for whatever (section,
    # subject) combinations it covers, fully replacing the report-derived
    # data for those keys. Each Batch List row already names its own
    # batch label, so the group is exactly (batch_label, faculty) — two
    # batches under the same faculty naturally become two groups.
    # canon_section() is applied here too so a Batch List section spelled
    # differently from the report's own Batch column (e.g. no dash before
    # the section letter) still lines up correctly.
    batch_map = {}  # (canon section, subject, roll) -> (batch_label, faculty)
    batch_subject_groups = {}  # (canon section, subject) -> set of (batch_label, faculty)
    if batch_entries:
        for section, subject, roll, batch_label, fac in batch_entries:
            key_section = canon_section(section)
            batch_map[(key_section, subject, roll)] = (batch_label, fac)
            batch_subject_groups.setdefault((key_section, subject), set()).add((batch_label, fac))
        for key, groups in batch_subject_groups.items():
            mapping[key] = set(groups)

    subject_groups = {k: sorted(v, key=group_sort_key) for k, v in mapping.items()}

    # ANY subject with more than one distinct GROUP in a section gets
    # split into a row per group — this isn't lab-only, and it isn't
    # faculty-only either. Languages and other theory subjects can
    # genuinely be taught in sub-groups by different faculty within the
    # same section (e.g. Hindi split into two groups), AND a single
    # faculty can cover two separate lab batches of the same subject —
    # both cases need their own row rather than being silently merged
    # into one combined number.
    multi_keys = {k for k, v in mapping.items() if len(v) > 1}

    student_group = {}
    for section, subj, roll, batch_label, fac in entries:
        key = (section, subj)
        if key in multi_keys and key not in batch_subject_groups:
            student_group[(section, subj, roll)] = (batch_label, fac)
    for (section, subj, roll), (batch_label, fac) in batch_map.items():
        if (section, subj) in multi_keys:
            student_group[(section, subj, roll)] = (batch_label, fac)

    return subject_groups, student_group


def write_faculty_map_sheet(wb, subject_groups):
    ws = wb.create_sheet(title=FACULTY_MAP_SHEET)
    ws.cell(row=1, column=1, value="Section")
    ws.cell(row=1, column=2, value="Subject")
    ws.cell(row=1, column=3, value="BatchLabel")
    ws.cell(row=1, column=4, value="Faculty")
    i = 2
    for (section, subj), groups in sorted(subject_groups.items()):
        for batch_label, fac in groups:
            ws.cell(row=i, column=1, value=section)
            ws.cell(row=i, column=2, value=subj)
            ws.cell(row=i, column=3, value=batch_label or "")
            ws.cell(row=i, column=4, value=fac)
            i += 1
    ws.sheet_state = "hidden"


def write_student_faculty_map_sheet(wb, student_group):
    ws = wb.create_sheet(title=STUDENT_FACULTY_MAP_SHEET)
    ws.cell(row=1, column=1, value="Section")
    ws.cell(row=1, column=2, value="Subject")
    ws.cell(row=1, column=3, value="Roll")
    ws.cell(row=1, column=4, value="BatchLabel")
    ws.cell(row=1, column=5, value="Faculty")
    for i, ((section, subj, roll), (batch_label, fac)) in enumerate(sorted(student_group.items()), start=2):
        ws.cell(row=i, column=1, value=section)
        ws.cell(row=i, column=2, value=subj)
        ws.cell(row=i, column=3, value=roll)
        ws.cell(row=i, column=4, value=batch_label or "")
        ws.cell(row=i, column=5, value=fac)
    ws.sheet_state = "hidden"


def resolve_groups_for_sheet(sheet_title, subject, subject_groups):
    """Sorted list of distinct (batch_label, faculty) GROUPS for `subject`
    on the debar sheet `sheet_title`.

    - Section-specific sheets ('CG 2025 A'): direct (section, subject)
      lookup — exactly the groups for that section. `subject_groups` is
      keyed by canon_section() (see its docstring), so `sheet_title` is
      canon'd here before the lookup too — otherwise a report sheet
      titled e.g. 'BCA AIML 2026 - A' would never match a Batch List
      entry spelled 'BCA AIML 2026 A' (no dash).
    - Series-aggregate sheets ('CG 2025 GEN' / 'CG 2025 GEN ALL'): these
      span every section in the series, so every distinct group across
      those sections is unioned in — if two sections use two different
      faculty (or the same faculty on two different batches) for the same
      subject, all of them show up, not just one.
    """
    direct = subject_groups.get((canon_section(sheet_title), subject))
    if direct is not None:
        return list(direct)

    base = sheet_title
    for suffix in (" GEN ALL", " GEN"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break
    else:
        return []  # not a GEN/GEN ALL sheet and no direct section match
    base = canon_section(base)

    groups = set()
    for (section, subj), grp_list in subject_groups.items():
        if subj == subject and canon_section(series_of(section)) == base:
            groups.update(grp_list)
    return sorted(groups, key=group_sort_key)


def in_subfolder(output_path, folder_name):
    """Redirect output_path into folder_name, creating it if needed.

    e.g. in_subfolder("VMS_Report.xlsx", "VMS Report")
         -> "VMS Report/VMS_Report.xlsx" (folder created)
    If output_path already includes a directory (e.g. "reports/out.xlsx"),
    the subfolder is nested inside that directory instead of the cwd.
    """
    parent, filename = os.path.split(output_path)
    target_dir = os.path.join(parent, folder_name) if parent else folder_name
    os.makedirs(target_dir, exist_ok=True)
    return os.path.join(target_dir, filename)


def build_report(input_path, output_path, low=0.0, high=75.0, dept="ALL", exclude=None, include=None, include_soft_skill=True, batch_list_path=None):
    output_path = in_subfolder(output_path, "VMS Report")
    extra_ignore = None if include_soft_skill else [SOFT_SKILL_KEYWORD]
    G, C = load_raw(input_path, extra_ignore)
    batch_entries = load_batch_list(batch_list_path) if batch_list_path else None
    subject_groups, student_group = build_faculty_maps(G, C, batch_entries)
    df = apply_subject_filters(G, C, exclude=exclude, include=include)
    if not df:
        raise ValueError("No data left after filtering — check the input file / include / exclude list.")

    all_depts = sorted({r["_dept"] for r in df})
    if dept and dept.strip().upper() not in ("ALL", ""):
        wanted = dept.strip().upper()
        depts = [d for d in all_depts if d == wanted]
        if not depts:
            raise ValueError(
                f"No rows found for department '{dept}'. Departments present in this file: "
                f"{', '.join(all_depts)}"
            )
    else:
        depts = all_depts

    wb = Workbook()
    wb.remove(wb.active)
    summaries = []

    for d in depts:
        d_rows = [r for r in df if r["_dept"] == d]
        for ser in get_series(d_rows, C):
            s_rows = series_df(d_rows, ser, C)
            s_subs = get_subjects(s_rows, C, extra_ignore)
            if not s_subs:
                continue

            g_grid, _ = process_grid(s_rows, s_subs, low, high, False, C)
            if g_grid:
                write_sheet(wb, f"{ser} GEN", g_grid, s_rows, s_subs, high, C)

            a_grid, _ = process_grid(s_rows, s_subs, low, high, True, C)
            if a_grid:
                write_sheet(wb, f"{ser} GEN ALL", a_grid, s_rows, s_subs, high, C)

            sections = sorted({str(r[C["batch"]]) for r in s_rows})
            for sec in sections:
                sec_rows = [r for r in s_rows if str(r[C["batch"]]) == sec]
                grid, _ = process_grid(sec_rows, s_subs, low, high, False, C)
                if grid:
                    summaries.append({"Section": sec, "Count": len(grid) - 1})
                    write_sheet(wb, sec.replace("/", "-"), grid, sec_rows, s_subs, high, C)

    if summaries:
        write_summary_sheet(wb, summaries)

    if not wb.sheetnames:
        raise ValueError("No sheets generated — check From/To % range and Department filter.")

    if subject_groups:
        write_faculty_map_sheet(wb, subject_groups)
    if student_group:
        write_student_faculty_map_sheet(wb, student_group)

    wb.save(output_path)
    return output_path, summaries


# ══════════════════════════════════════════════════════════════════════
# PART 1B — TENTATIVE DEBAR LIST FORMATTER (ported from the HTML tool)
# ══════════════════════════════════════════════════════════════════════
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "presidency_logo.jpg")

# Row heights (points) and column widths, taken from the reference template
DEBAR_ROW1_H, DEBAR_ROW2_H, DEBAR_ROW3_H = 54.6, 23.4, 72
DEBAR_DATA_ROW_H, DEBAR_FOOTER_H, DEBAR_SIG1_H = 18, 18, 29.4
DEBAR_COL_A_W, DEBAR_COL_B_W, DEBAR_COL_C_W, DEBAR_COL_SUBJ_W = 5.9, 9.1, 33, 8.5
DEBAR_COL_SIG_W = 16  # wider than a subject column — room for an actual signature

# Displayed logo size in pixels, taken from the reference drawing (EMU / 9525)
DEBAR_LOGO_W = 8801100 / 9525
DEBAR_LOGO_H = 563880 / 9525

DEBAR_THIN = Side(style="thin")
DEBAR_BORDER_ALL = Border(left=DEBAR_THIN, right=DEBAR_THIN, top=DEBAR_THIN, bottom=DEBAR_THIN)


def excel_col_width_to_px(width):
    """Excel's stored column width (in 'characters', the number you see in
    the column-width dialog) to on-screen pixels — the standard Calibri-11
    approximation Excel itself uses (7px per character plus 5px padding).
    Only needs to be roughly right: used to keep the notice-board logo
    from visually overflowing past the table's own right edge, not for
    pixel-perfect layout."""
    return round(width * 7 + 5)


def _dfont(size, bold=False):
    return Font(name="Calibri", size=size, bold=bold)


def build_debar_style(scale=1.0):
    """Build the debar-list cell style set, with every font size scaled by
    `scale` (rounded to the nearest point). scale=1.0 reproduces the
    original sizes exactly; a bigger scale is used for the notice-board
    variant so it reads clearly from a distance."""
    def sz(n):
        return max(1, round(n * scale))
    return {
        "title": dict(font=_dfont(sz(14), True), alignment=Alignment(horizontal="center"), border=DEBAR_BORDER_ALL),
        "hdrNarrow": dict(font=_dfont(sz(11), True), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "hdrSubject": dict(font=_dfont(sz(9), True), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=DEBAR_BORDER_ALL),
        "dataSlNo": dict(font=_dfont(sz(11)), alignment=Alignment(horizontal="center"), border=DEBAR_BORDER_ALL),
        "dataRoll": dict(font=_dfont(sz(11)), alignment=Alignment(), border=DEBAR_BORDER_ALL),
        "dataName": dict(font=_dfont(sz(11)), alignment=Alignment(), border=DEBAR_BORDER_ALL),
        "dataSubject": dict(font=_dfont(sz(11)), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "dataSubjectBold": dict(font=_dfont(sz(11), True), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "footerLabel": dict(font=_dfont(sz(11), True), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "footerCount": dict(font=_dfont(sz(11)), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "sigLabel": dict(font=_dfont(sz(11), True), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
        "sigClass": dict(font=_dfont(sz(11), True), alignment=Alignment(), border=None),
        "sigHod": dict(font=_dfont(sz(11), True), alignment=Alignment(horizontal="center", vertical="center"), border=None),
        "legendHdr": dict(font=Font(name="Calibri", size=sz(11), bold=True, color="FFFFFF"),
                           alignment=Alignment(horizontal="center", vertical="center"),
                           border=DEBAR_BORDER_ALL, fill=HEADER_FILL),
        "legendSubject": dict(font=_dfont(sz(9), True), alignment=Alignment(horizontal="center", vertical="center", wrap_text=True), border=DEBAR_BORDER_ALL),
        "legendFaculty": dict(font=_dfont(sz(11)), alignment=Alignment(horizontal="center", vertical="center"), border=DEBAR_BORDER_ALL),
    }


DEBAR_STYLE = build_debar_style(1.0)

# Legend table column zones (matching the reference format). SUBJECT spans
# the same three columns as Sl No/Roll No/Student Name in the debar table
# above it, so it gets enough width to hold long subject names cleanly
# instead of being squeezed into the narrow "Sl No." column alone.
LEGEND_SUBJECT_START, LEGEND_SUBJECT_END = 1, 3
LEGEND_FACULTY_START, LEGEND_FACULTY_END = 4, 6
LEGEND_COUNT_START, LEGEND_COUNT_END = 7, 8
LEGEND_PLAN_START, LEGEND_PLAN_END = 9, 13
LEGEND_SIG_START, LEGEND_SIG_END = 14, 15
LEGEND_ROW_H = 30


def _dapply(cell, style_name, styles=None):
    st = (styles or DEBAR_STYLE)[style_name]
    cell.font = st["font"]
    cell.alignment = st["alignment"]
    if st["border"] is not None:
        cell.border = st["border"]
    if st.get("fill") is not None:
        cell.fill = st["fill"]


def _apply_border_range(ows, r1, c1, r2, c2, border):
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            ows.cell(row=rr, column=cc).border = border


def _write_debar_legend(ows, start_row, legend_entries):
    """Write the SUBJECT / FACULTY IN CHARGE / NO of students / Plan of
    Action / Signature legend table just below the tentative debar list.

    `legend_entries` is a list of {"subject": str, "rows": [{"faculty": str,
    "count": int}, ...]}. A subject normally gets exactly one row. A lab
    subject taught by more than one faculty (different lab batches in the
    same section) gets one row PER faculty, each with that faculty's own
    student count — and the SUBJECT cell is merged vertically across those
    rows so it still reads as one subject with two teachers underneath it.

    Returns the row number of the last legend row written.
    """
    header_row = start_row
    zones = [
        (LEGEND_SUBJECT_START, LEGEND_SUBJECT_END, "SUBJECT"),
        (LEGEND_FACULTY_START, LEGEND_FACULTY_END, "FACULTY IN CHARGE"),
        (LEGEND_COUNT_START, LEGEND_COUNT_END, "NO of students"),
        (LEGEND_PLAN_START, LEGEND_PLAN_END, "Plan of Action"),
        (LEGEND_SIG_START, LEGEND_SIG_END, "Signature"),
    ]
    for c1, c2, label in zones:
        if c2 > c1:
            ows.merge_cells(start_row=header_row, start_column=c1, end_row=header_row, end_column=c2)
        cell = ows.cell(row=header_row, column=c1, value=label)
        _dapply(cell, "legendHdr")
        _apply_border_range(ows, header_row, c1, header_row, c2, DEBAR_BORDER_ALL)
        # fill covers the whole merged zone, not just the anchor cell
        for c in range(c1, c2 + 1):
            ows.cell(row=header_row, column=c).fill = HEADER_FILL

    r = header_row + 1
    for entry in legend_entries:
        subj = entry["subject"]
        sub_rows = entry["rows"] or [{"faculty": "", "count": 0}]
        first_r = r

        for sr in sub_rows:
            ows.row_dimensions[r].height = LEGEND_ROW_H
            for c1, c2 in ((LEGEND_FACULTY_START, LEGEND_FACULTY_END),
                           (LEGEND_COUNT_START, LEGEND_COUNT_END),
                           (LEGEND_PLAN_START, LEGEND_PLAN_END),
                           (LEGEND_SIG_START, LEGEND_SIG_END)):
                ows.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

            faculty_cell = ows.cell(row=r, column=LEGEND_FACULTY_START, value=sr["faculty"])
            _dapply(faculty_cell, "legendFaculty")

            count_cell = ows.cell(row=r, column=LEGEND_COUNT_START, value=sr["count"])
            _dapply(count_cell, "dataSubjectBold")

            # Plan of Action / Signature: left blank for the department to fill in
            _dapply(ows.cell(row=r, column=LEGEND_PLAN_START), "legendFaculty")
            _dapply(ows.cell(row=r, column=LEGEND_SIG_START), "legendFaculty")

            _apply_border_range(ows, r, LEGEND_FACULTY_START, r, LEGEND_SIG_END, DEBAR_BORDER_ALL)
            r += 1

        last_r = r - 1
        # SUBJECT cell: merged across every sub-row this subject got
        # (a single-row subject just gets a 1-row "merge", harmless).
        ows.merge_cells(start_row=first_r, start_column=LEGEND_SUBJECT_START,
                         end_row=last_r, end_column=LEGEND_SUBJECT_END)
        subj_cell = ows.cell(row=first_r, column=LEGEND_SUBJECT_START, value=subj)
        _dapply(subj_cell, "legendSubject")
        _apply_border_range(ows, first_r, LEGEND_SUBJECT_START, last_r, LEGEND_SUBJECT_END, DEBAR_BORDER_ALL)

    return r - 1


def find_latest_vms_report(near_path=None):
    """Return the most recently modified .xlsx in the "VMS Report" folder
    (searched next to near_path's directory if given, else the cwd)."""
    parent = os.path.dirname(near_path) if near_path else ""
    folder = os.path.join(parent, "VMS Report") if parent else "VMS Report"
    candidates = glob.glob(os.path.join(folder, "*.xlsx"))
    if not candidates:
        raise FileNotFoundError(
            f"No VMS Report files found in '{folder}'. Run `format` (or `pipeline`) first."
        )
    return max(candidates, key=os.path.getmtime)


def resolve_input_path(input_path):
    """Resolve the debar step's input file, always preferring an exact
    match if one exists, otherwise auto-picking the most recently
    modified file in the "VMS Report" folder — so you never have to type
    an exact (often timestamped) filename."""
    if input_path is None:
        latest = find_latest_vms_report()
        print(f"  (no input given — using latest: {latest})")
        return latest
    if os.path.exists(input_path):
        return input_path
    parent, filename = os.path.split(input_path)
    candidate = os.path.join(parent, "VMS Report", filename) if parent else os.path.join("VMS Report", filename)
    if os.path.exists(candidate):
        print(f"  (found input at: {candidate})")
        return candidate
    latest = find_latest_vms_report(input_path)
    print(f"  ('{input_path}' not found — using latest instead: {latest})")
    return latest


ABSTRACT_COL_SECTION_W = 16
ABSTRACT_COL_SUBJECT_W = 34
ABSTRACT_COL_FACULTY_W = 26
ABSTRACT_COL_BUCKET_W = 12
ABSTRACT_COL_TOTAL_W = 9
ABSTRACT_ROW_H = 22
ABSTRACT_ROW1_H, ABSTRACT_ROW2_H = 54.6, 23.4  # logo row, title row — same as the debar list


def _abstract_style(bold=False, fill=None, wrap=False):
    return dict(
        font=Font(name="Calibri", size=11, bold=bold, color="FFFFFF" if fill else "000000"),
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=wrap),
        border=BORDER,
        fill=fill,
    )


def _abstract_apply(cell, style):
    cell.font = style["font"]
    cell.alignment = style["alignment"]
    cell.border = style["border"]
    if style["fill"] is not None:
        cell.fill = style["fill"]


def _abstract_section_rows(section, subject_rows, C, batch_map):
    """For one section, one row per subject (or one per GROUP — a
    (batch_label, faculty) pair — if a subject has more than one distinct
    group), each as {"subject":..., "faculty":..., "buckets":[6 ints],
    "total": int}. When a subject splits across more than one group, the
    faculty display is prefixed with its lab batch label ('Batch 1 -
    NAME'), same as the debar legend. This also correctly separates two
    lab batches taught by the SAME faculty — they're still two different
    groups and get two different rows, rather than being merged into one
    combined count. A single-group subject just shows the plain name, no
    label needed."""
    subjects = get_subjects(subject_rows, C)
    out_rows = []
    key_section = canon_section(section)  # batch_map is keyed by canon_section() — see its docstring
    for subj in subjects:
        srows = [r for r in subject_rows if r[C["subject"]] == subj]
        by_group = {}  # (batch_label, faculty) -> [6 bucket counts]
        all_groups = set()  # every distinct group for this subject/section, even with zero shortage count
        for r in srows:
            roll = str(r.get(C["roll"]) or "").strip()
            pct = r.get(C["attendance"])
            bidx = bucket_index(pct)
            entry = batch_map.get((key_section, subj, roll))
            if entry:
                batch_label, fac = entry
            else:
                batch_label, fac = None, str(r.get(C["faculty"]) or "").strip()
            fac = fac or "—"
            group_key = (batch_label, fac)
            all_groups.add(group_key)
            if bidx is None:
                continue
            by_group.setdefault(group_key, [0] * 6)[bidx] += 1

        if not by_group:
            # Nobody fell into a shortage bucket (e.g. all-zero count for
            # this subject) — still show the group(s) rather than leaving
            # the column blank. Show the batch label whenever it's known,
            # even if this subject only has ONE group in this section —
            # a single batch is still a batch and shouldn't be hidden.
            if all_groups:
                for batch_label, fac in sorted(all_groups, key=group_sort_key):
                    display = group_label(batch_label, fac) if batch_label else fac
                    out_rows.append({"subject": subj, "faculty": display, "buckets": [0] * 6, "total": 0})
            else:
                out_rows.append({"subject": subj, "faculty": "", "buckets": [0] * 6, "total": 0})
            continue

        for batch_label, fac in sorted(by_group, key=group_sort_key):
            buckets = by_group[(batch_label, fac)]
            display = group_label(batch_label, fac) if batch_label else fac
            out_rows.append({"subject": subj, "faculty": display, "buckets": buckets, "total": sum(buckets)})
    return out_rows


def write_abstract_sheet(ws, sections, G, C, batch_map, date_str=None, program="BCA", sem_heading=""):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = ABSTRACT_COL_SECTION_W
    ws.column_dimensions["B"].width = ABSTRACT_COL_SUBJECT_W
    ws.column_dimensions["C"].width = ABSTRACT_COL_FACULTY_W
    for i in range(6):
        ws.column_dimensions[get_column_letter(4 + i)].width = ABSTRACT_COL_BUCKET_W
    ws.column_dimensions[get_column_letter(10)].width = ABSTRACT_COL_TOTAL_W

    hdr_style = _abstract_style(bold=True, fill=HEADER_FILL)
    subj_style = _abstract_style(bold=True, wrap=True)
    fac_style = _abstract_style(bold=False)
    count_style = _abstract_style(bold=False)
    total_style = _abstract_style(bold=True)
    sec_style = _abstract_style(bold=True, wrap=True)
    title_style = _abstract_style(bold=True)

    last_col_letter = get_column_letter(10)

    # Row 1: Presidency logo
    ws.row_dimensions[1].height = ABSTRACT_ROW1_H
    if os.path.exists(LOGO_PATH):
        img = XLImage(LOGO_PATH)
        img.width, img.height = DEBAR_LOGO_W, DEBAR_LOGO_H
        ws.add_image(img, "A1")

    # Row 2: heading, e.g. "I SEM BCA ABSTRACT AS OF 07.09.2026"
    ws.row_dimensions[2].height = ABSTRACT_ROW2_H
    ws.merge_cells(f"A2:{last_col_letter}2")
    title_cell = ws.cell(row=2, column=1)
    title_cell.value = f"{sem_heading} SEM {program} ABSTRACT AS OF {date_str}".strip()
    _abstract_apply(title_cell, title_style)

    r = 3
    for section in sections:
        sec_rows = [row for row in G if str(row.get(C["batch"]) or "").strip().replace("/", "-") == section]
        rows_for_section = _abstract_section_rows(section, sec_rows, C, batch_map)
        if not rows_for_section:
            continue

        header_row = r
        headers = ["SECTION", "Subject", "FACULTY INCHARGE"] + ABSTRACT_BUCKET_LABELS + ["Total"]
        for c, label in enumerate(headers, start=1):
            cell = ws.cell(row=header_row, column=c, value=label)
            _abstract_apply(cell, hdr_style)
        ws.row_dimensions[header_row].height = ABSTRACT_ROW_H
        r += 1

        section_first_row = r
        subj_first_row = {}
        for entry in rows_for_section:
            ws.row_dimensions[r].height = ABSTRACT_ROW_H
            if entry["subject"] not in subj_first_row:
                subj_first_row[entry["subject"]] = r
            fcell = ws.cell(row=r, column=3, value=entry["faculty"])
            _abstract_apply(fcell, fac_style)
            for i, v in enumerate(entry["buckets"]):
                ccell = ws.cell(row=r, column=4 + i, value=v)
                _abstract_apply(ccell, count_style)
            tcell = ws.cell(row=r, column=10, value=entry["total"])
            _abstract_apply(tcell, total_style)
            r += 1
        section_last_row = r - 1

        # merge SUBJECT vertically for subjects with more than one faculty row
        i = section_first_row
        while i <= section_last_row:
            subj = rows_for_section[i - section_first_row]["subject"]
            j = i
            while j + 1 <= section_last_row and rows_for_section[j + 1 - section_first_row]["subject"] == subj:
                j += 1
            ws.merge_cells(start_row=i, start_column=2, end_row=j, end_column=2)
            scell = ws.cell(row=i, column=2, value=subj)
            _abstract_apply(scell, subj_style)
            for rr in range(i, j + 1):
                ws.cell(row=rr, column=2).border = BORDER
            i = j + 1

        # merge SECTION for the whole block
        ws.merge_cells(start_row=section_first_row, start_column=1, end_row=section_last_row, end_column=1)
        secell = ws.cell(row=section_first_row, column=1, value=section)
        _abstract_apply(secell, sec_style)
        for rr in range(section_first_row, section_last_row + 1):
            ws.cell(row=rr, column=1).border = BORDER

        r = section_last_row + 3  # 2 blank rows between section blocks


def build_abstract(input_path, batch_list_path, output_path, date_str=None, program="BCA"):
    """Per-semester abstract workbook: for every section (that the Lab
    Batch List covers), every subject gets a row (or one row per faculty,
    for a lab/internship taught by more than one faculty) showing FACULTY
    INCHARGE and a 6-band breakdown of the shortage population by
    attendance % (0-29.99 ... 70-74.99), plus a Total. Faculty for labs
    and internship comes from the Batch List (cross-referenced by Reg No);
    theory subjects use the consolidated report's own Staff Name column.
    Each sheet gets the Presidency logo and a heading like "I SEM BCA
    ABSTRACT AS OF <date>" at the top.

    Soft Skill is ALWAYS excluded from the abstract, regardless of the
    Soft Skill choice made for the main VMS report.
    """
    if date_str is None:
        date_str = time.strftime("%d.%m.%Y")
    output_path = in_subfolder(output_path, "Abstract Report")
    # Soft Skill is never part of the abstract.
    G, C = load_raw(input_path, [SOFT_SKILL_KEYWORD])
    batch_entries = load_batch_list(batch_list_path)
    if not batch_entries:
        raise ValueError("No usable rows found in the Batch List workbook — check its column headers match "
                          "'Reg No.', 'Batch', 'Course Community Name', 'Faculty'.")

    # canon_section() so a Batch List section spelled without the dash
    # ('BCA AIML 2026 A') still matches the report's own spelling
    # ('BCA AIML 2026 - A') — see canon_section()'s docstring. batch_map
    # is looked up the same way inside _abstract_section_rows().
    batch_map = {}
    covered_sections = set()
    for section, subject, roll, batch_label, fac in batch_entries:
        key_section = canon_section(section)
        batch_map[(key_section, subject, roll)] = (batch_label, fac)
        covered_sections.add(key_section)

    section_sem = {}
    for r in G:
        sec = str(r.get(C["batch"]) or "").strip().replace("/", "-")
        if canon_section(sec) in covered_sections and sec not in section_sem:
            section_sem[sec] = str(r.get(C["sem"]) or "").strip()

    missing = covered_sections - {canon_section(s) for s in section_sem}
    if missing:
        print(f"  (note: {len(missing)} Batch List section(s) not found in the consolidated report, skipped: {sorted(missing)})")

    if not section_sem:
        raise ValueError("None of the Batch List's sections were found in the consolidated report — "
                          "check both files are for the same term.")

    sems_present = sorted(set(section_sem.values()), key=lambda s: int(re.sub(r"\D", "", s) or 0))

    wb = Workbook()
    wb.remove(wb.active)
    for sem in sems_present:
        sections = sorted(s for s, sv in section_sem.items() if sv == sem)
        sheet_name = f"SUB {sem_label(sem)} SEM"[:31]
        ws = wb.create_sheet(sheet_name)
        write_abstract_sheet(ws, sections, G, C, batch_map, date_str=date_str, program=program,
                              sem_heading=sem_label(sem))

    if not wb.sheetnames:
        raise ValueError("No abstract sheets generated.")

    wb.save(output_path)
    return output_path, sems_present


def build_debar_list(input_path, output_path, date_str=None, notice_board=False,
                      font_scale=1.6, row_scale=1.6):
    """
    For every sheet (batch) in a raw VMS export, build a matching
    "TENTATIVE DEBAR LIST I" sheet: college header image, merged title row,
    the same fonts/borders/column widths as the reference layout, subject
    columns kept as-is, per-subject shortage counts at the bottom, and
    Subject Teacher / Class Teacher / HOD signature lines.

    notice_board=True builds a print-friendly variant meant for posting on
    a physical notice board instead of internal/faculty use: the trailing
    Student Signature column, the SUBJECT / FACULTY IN CHARGE / NO of
    students / Plan of Action / Signature legend table, the per-subject
    shortage-count footer row, and the Class Teacher / HOD / Principal
    signature line are all left out entirely (none of that is useful to a
    student reading the board) — the sheet ends right after the student
    data. Every font is scaled up by `font_scale` and every row height by
    `row_scale` for legibility from a distance; the Sl No / Roll No /
    Student Name columns are widened by the same factor so they don't
    look cramped at the bigger font size, and the college logo is
    enlarged by the row-height factor. No font is ever scaled DOWN to
    help the page fit — including the heading — page-fit is handled
    purely by print scaling (fitToWidth=1, fitToHeight=0 below), so a
    section with more students simply flows onto more printed pages
    landscape rather than the heading or data being shrunk to squeeze
    onto one page.
    """
    if date_str is None:
        date_str = time.strftime("%d.%m.%Y")

    input_path = resolve_input_path(input_path)
    styles = build_debar_style(font_scale) if notice_board else DEBAR_STYLE
    row2_h = DEBAR_ROW2_H * row_scale if notice_board else DEBAR_ROW2_H
    row3_h = DEBAR_ROW3_H * row_scale if notice_board else DEBAR_ROW3_H
    data_row_h = DEBAR_DATA_ROW_H * row_scale if notice_board else DEBAR_DATA_ROW_H
    folder_name = "Tentative Debar List (Notice Board)" if notice_board else "Tentative Debar List"
    output_path = in_subfolder(output_path, folder_name)

    src = load_workbook(input_path, data_only=True)
    out = Workbook()
    out.remove(out.active)  # drop the default blank sheet

    # Subject -> group(s) in charge, read from column Q of the consolidated
    # report earlier by `format`/`pipeline` and carried here in hidden
    # sheets of the VMS Report workbook. A group is (batch_label, faculty)
    # — two lab batches taught by the same faculty are two distinct groups.
    subject_groups = {}  # (section, subject) -> list of (batch_label, faculty)
    if FACULTY_MAP_SHEET in src.sheetnames:
        fm = src[FACULTY_MAP_SHEET]
        for row in fm.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1] and row[3]:
                section = str(row[0]).strip()
                subj = str(row[1]).strip()
                batch_label = str(row[2]).strip() if row[2] else None
                fac = str(row[3]).strip()
                subject_groups.setdefault((section, subj), []).append((batch_label, fac))

    student_group = {}  # (section, subject, roll) -> (batch_label, faculty), multi-group subjects only
    if STUDENT_FACULTY_MAP_SHEET in src.sheetnames:
        sfm = src[STUDENT_FACULTY_MAP_SHEET]
        for row in sfm.iter_rows(min_row=2, values_only=True):
            if row and row[0] and row[1] and row[2]:
                section = str(row[0]).strip()
                subj = str(row[1]).strip()
                roll = str(row[2]).strip()
                batch_label = str(row[3]).strip() if row[3] else None
                fac = str(row[4]).strip() if row[4] else ""
                student_group[(section, subj, roll)] = (batch_label, fac)

    log = []  # list of (sheet_name, ok, detail)
    sheets_done = 0

    for sws in src.worksheets:
        if sws.title in (FACULTY_MAP_SHEET, STUDENT_FACULTY_MAP_SHEET):
            continue
        try:
            headers = [cell.value for cell in sws[1]]

            def find_idx(name):
                return headers.index(name) if name in headers else -1

            idx_theory = find_idx("Theory Avg")
            idx_final = find_idx("Final Avg")
            idx_range = find_idx("Subjects in Range")

            if idx_theory == -1 or idx_final == -1 or idx_range == -1:
                log.append((sws.title, False, "Missing 'Theory Avg' / 'Final Avg' / 'Subjects in Range' column — sheet skipped."))
                continue

            # subject headers sit in columns F.. up to (not incl.) Theory Avg
            # -> 0-based slice(5, idx_theory)
            subject_headers = headers[5:idx_theory]
            n_subjects = len(subject_headers)
            # Notice-board copies drop the trailing Student Signature
            # column entirely — nobody signs a sheet pinned to a board —
            # so n_cols has no +1 for it in that case.
            n_cols = 3 + n_subjects + 2 + (0 if notice_board else 1)
            final_avg_col = 3 + n_subjects + 1
            shortage_col = 3 + n_subjects + 2
            signature_col = None if notice_board else 3 + n_subjects + 3
            last_col_letter = get_column_letter(n_cols)

            name = sws.title[:31]
            ows = out.create_sheet(name)

            # Sl No / Roll No / Student Name columns are widened for the
            # notice-board copy (same scale as the enlarged font) so they
            # stay clearly readable rather than cramped/truncated at the
            # bigger font size.
            col_scale = font_scale if notice_board else 1.0
            ows.column_dimensions["A"].width = DEBAR_COL_A_W * col_scale
            ows.column_dimensions["B"].width = DEBAR_COL_B_W * col_scale
            ows.column_dimensions["C"].width = DEBAR_COL_C_W * col_scale
            for c in range(4, n_cols + 1):
                ows.column_dimensions[get_column_letter(c)].width = DEBAR_COL_SUBJ_W
            if signature_col is not None:
                ows.column_dimensions[get_column_letter(signature_col)].width = DEBAR_COL_SIG_W

            # Row 1: logo — "dragged" bigger for the notice-board copy
            # (same enlargement factor as the row heights) so it's easy
            # to spot from a distance on a physical board — but never
            # wider than the table itself: for sheets with fewer subject
            # columns (e.g. a GEN sheet), the enlarged logo could
            # otherwise overhang past the table's right edge, so its
            # width is capped to the table's total column width (aspect
            # ratio preserved, so it's never stretched/squashed either).
            logo_scale = row_scale if notice_board else 1.0
            ows.row_dimensions[1].height = DEBAR_ROW1_H * logo_scale
            if os.path.exists(LOGO_PATH):
                img = XLImage(LOGO_PATH)
                logo_w, logo_h = DEBAR_LOGO_W * logo_scale, DEBAR_LOGO_H * logo_scale
                if notice_board:
                    table_width_px = sum(
                        excel_col_width_to_px(ows.column_dimensions[get_column_letter(c)].width)
                        for c in range(1, n_cols + 1)
                    )
                    if logo_w > table_width_px:
                        shrink = table_width_px / logo_w
                        logo_w, logo_h = logo_w * shrink, logo_h * shrink
                img.width, img.height = logo_w, logo_h
                ows.add_image(img, "A1")

            # Row 2: title — font size for this row comes from build_debar_style()
            # via `styles["title"]`, which for the notice-board copy is
            # SCALED UP by font_scale (never reduced) along with every
            # other font on the sheet, and is never separately shrunk to
            # help the page fit — page-fit is handled entirely by
            # fitToWidth/fitToHeight below, not by touching font sizes.
            ows.row_dimensions[2].height = row2_h
            ows.merge_cells(f"A2:{last_col_letter}2")
            title_cell = ows.cell(row=2, column=1)
            title_cell.value = f'TENTATIVE DEBAR LIST I - "{sws.title}" as of {date_str}' + \
                (" (NOTICE BOARD COPY)" if notice_board else "")
            _dapply(title_cell, "title", styles)
            for c in range(1, n_cols + 1):
                ows.cell(row=2, column=c).border = DEBAR_BORDER_ALL

            # Row 3: header
            ows.row_dimensions[3].height = row3_h
            header_values = ["Sl No.", "Roll No.", "Student Name"] + list(subject_headers) + \
                             ["Final Avg", "No of subjects having Shortage"] + \
                             ([] if notice_board else ["Student Signature"])
            for i, val in enumerate(header_values):
                c = i + 1
                cell = ows.cell(row=3, column=c)
                cell.value = val
                _dapply(cell, "hdrNarrow" if c <= 3 else "hdrSubject", styles)

            # Data rows
            counts = [0] * n_subjects
            multi_group_counts = [Counter() for _ in range(n_subjects)]  # group_key -> count
            row_out = 4
            for srow in sws.iter_rows(min_row=2):
                student_name = srow[2].value if len(srow) > 2 else None
                sl_no = srow[0].value if len(srow) > 0 else None
                if student_name is None or str(student_name).strip() == "":
                    continue
                if not isinstance(sl_no, (int, float)):
                    continue

                roll = srow[1].value
                row_batch = srow[3].value if len(srow) > 3 else None
                row_section = str(row_batch).strip().replace("/", "-") if row_batch not in (None, "") else sws.title
                subj_vals = [srow[5 + i].value if len(srow) > 5 + i else None for i in range(n_subjects)]
                final_avg = srow[idx_final].value if len(srow) > idx_final else None
                shortage_n = srow[idx_range].value if len(srow) > idx_range else None

                ows.row_dimensions[row_out].height = data_row_h
                c1 = ows.cell(row=row_out, column=1, value=sl_no); _dapply(c1, "dataSlNo", styles)
                c2 = ows.cell(row=row_out, column=2, value=roll); _dapply(c2, "dataRoll", styles)
                c3 = ows.cell(row=row_out, column=3, value=student_name); _dapply(c3, "dataName", styles)
                for i, v in enumerate(subj_vals):
                    cell = ows.cell(row=row_out, column=4 + i, value=(None if v == "" else v))
                    _dapply(cell, "dataSubject", styles)
                    if v not in (None, ""):
                        counts[i] += 1
                        grp = student_group.get((canon_section(row_section), str(subject_headers[i]).strip(), str(roll).strip()))
                        if grp and grp[1]:
                            multi_group_counts[i][grp] += 1
                cf = ows.cell(row=row_out, column=final_avg_col, value=final_avg); _dapply(cf, "dataSubject", styles)
                cs = ows.cell(row=row_out, column=shortage_col, value=shortage_n); _dapply(cs, "dataSubjectBold", styles)
                if signature_col is not None:
                    csig = ows.cell(row=row_out, column=signature_col); _dapply(csig, "dataSubject", styles)  # left blank for the student to sign
                row_out += 1

            last_data_row = row_out - 1

            if not notice_board:
                # Footer: counts per subject
                footer_row = last_data_row + 1
                ows.row_dimensions[footer_row].height = DEBAR_FOOTER_H
                for c in range(1, n_cols + 1):
                    cell = ows.cell(row=footer_row, column=c)
                    if c == 3:
                        cell.value = "No of students having shortage"
                        _dapply(cell, "footerLabel")
                    elif 4 <= c <= 3 + n_subjects:
                        cell.value = counts[c - 4]
                        _dapply(cell, "footerLabel")
                    else:
                        _dapply(cell, "footerCount")

                # Legend: SUBJECT / FACULTY IN CHARGE / NO of students /
                # Plan of Action / Signature — one row per subject column above.
                legend_entries = []
                for i, subj in enumerate(subject_headers):
                    subj_s = str(subj).strip()
                    grp_list = resolve_groups_for_sheet(sws.title, subj_s, subject_groups)

                    if len(grp_list) > 1 and multi_group_counts[i]:
                        # more than one distinct GROUP genuinely teaches this
                        # subject within this section — different lab batches
                        # (even under the same faculty), or a language/theory
                        # subject split into sub-groups by different faculty —
                        # one legend row per group, each with that group's own
                        # (accurate) student count and, when known (labs only),
                        # which batch it covers.
                        rows_for_subj = [
                            {"faculty": group_label(batch_label, fac),
                             "count": multi_group_counts[i].get((batch_label, fac), 0)}
                            for batch_label, fac in grp_list
                        ]
                    else:
                        names = " / ".join(dict.fromkeys(fac for _, fac in grp_list))
                        rows_for_subj = [{"faculty": names, "count": counts[i]}]

                    legend_entries.append({"subject": subj, "rows": rows_for_subj})

                legend_start = footer_row + 3
                last_legend_row = _write_debar_legend(ows, legend_start, legend_entries)
                sig_row2 = last_legend_row + 6

                # Class Teacher / HOD / Principal signature line — internal/
                # faculty copy only. Left out of the notice-board copy
                # entirely per request: nothing below the student data
                # except what's already there.
                class_cell = ows.cell(row=sig_row2, column=LEGEND_SUBJECT_START, value="Class Teacher Signature")
                _dapply(class_cell, "sigClass", styles)
                hod_cell = ows.cell(row=sig_row2, column=LEGEND_COUNT_START, value="HOD")
                _dapply(hod_cell, "sigHod", styles)
                principal_cell = ows.cell(row=sig_row2, column=LEGEND_PLAN_START, value="PRINCIPAL")
                _dapply(principal_cell, "sigHod", styles)

            ows.sheet_view.showGridLines = False
            ows.page_setup.orientation = "landscape"
            ows.page_setup.fitToPage = True
            ows.sheet_properties.pageSetUpPr.fitToPage = True
            ows.page_setup.fitToWidth = 1
            ows.page_setup.fitToHeight = 0
            ows.page_margins.left = 0.7
            ows.page_margins.right = 0.7
            ows.page_margins.top = 0.75
            ows.page_margins.bottom = 0.75
            ows.page_margins.header = 0.3
            ows.page_margins.footer = 0.3

            sheets_done += 1
            log.append((sws.title, True, f"{last_data_row - 3} students · {n_subjects} subject columns"))
        except Exception as err:
            log.append((sws.title, False, str(err)))

    if sheets_done == 0:
        raise ValueError("No sheets could be processed — check that the input file has the expected columns.")

    out.save(output_path)
    return output_path, log


# ══════════════════════════════════════════════════════════════════════
# PART 2 — LINWAYS DOWNLOAD (Selenium)
# ══════════════════════════════════════════════════════════════════════
def _by(kind):
    """Lazily resolve a Selenium By constant from its string name."""
    from selenium.webdriver.common.by import By
    return getattr(By, kind)


def build_driver(headless=True):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")  # so we can dismiss popups ourselves, not have Chrome swallow them
    opts.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
        # 2 = block. Stops the native "Site wants to show notifications" prompt
        # from ever appearing, so it can never block a click.
        "profile.default_content_setting_values.notifications": 2,
    })
    return webdriver.Chrome(options=opts)


def login(driver, username, password):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 20)

    kind, val = SELECTORS["username"]
    user_field = wait.until(EC.presence_of_element_located((_by(kind), val)))
    user_field.clear()
    user_field.send_keys(username)

    kind, val = SELECTORS["password"]
    pass_field = driver.find_element(_by(kind), val)
    pass_field.clear()
    pass_field.send_keys(password)

    kind, val = SELECTORS["login_button"]
    driver.find_element(_by(kind), val).click()

    wait.until(lambda d: "login" not in d.current_url.lower())
    print("Logged in. Current URL:", driver.current_url)


def dismiss_popups(driver):
    """Best-effort close of common modal/notice popups that may appear after login."""
    from selenium.webdriver.common.by import By
    common_close_selectors = [
        (By.CSS_SELECTOR, "button.close"),
        (By.CSS_SELECTOR, "[aria-label='Close']"),
        (By.XPATH, "//button[contains(text(),'Close')]"),
        (By.XPATH, "//button[contains(text(),'OK')]"),
        (By.XPATH, "//button[contains(text(),'Got it')]"),
        (By.XPATH, "//span[contains(@class,'close')]"),
        (By.CSS_SELECTOR, ".modal.show .close"),
        (By.CSS_SELECTOR, ".swal2-close"),
        (By.CSS_SELECTOR, ".swal2-confirm"),
    ]
    for kind, val in common_close_selectors:
        try:
            elems = driver.find_elements(kind, val)
            for el in elems:
                if el.is_displayed():
                    el.click()
                    time.sleep(0.5)
        except Exception:
            pass


def set_date_range(driver, from_date, to_date):
    """Fill in the report page's date range field(s) before clicking Search.

    from_date/to_date are plain strings (e.g. "01/06/2026") passed straight
    through to send_keys — match whatever format the real field expects.
    No-op if both are None, so existing behavior is unchanged when the
    caller doesn't ask for a date filter.

    NOTE: SELECTORS["date_from"] / SELECTORS["date_to"] are placeholders
    (see the SELECTOR CONFIG section and instructions at the bottom of this
    file) — confirm the real selectors via DevTools before relying on this.
    """
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    if not from_date and not to_date:
        return

    wait = WebDriverWait(driver, 20)

    def fill(step_name, value):
        if not value:
            return
        kind, val = SELECTORS[step_name]
        el = wait.until(EC.element_to_be_clickable((_by(kind), val)))
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", el
        )
        el.click()
        try:
            el.clear()
        except Exception:
            pass  # some datepicker widgets reject .clear() on a readonly input — ignore and just type
        el.send_keys(value)
        time.sleep(0.3)

    fill("date_from", from_date)
    fill("date_to", to_date)
    dismiss_popups(driver)


def download_consolidated_report(driver, from_date=None, to_date=None):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, 20)
    debug_dir = os.path.join(DOWNLOAD_DIR, "..", "debug_screenshots")
    debug_dir = os.path.abspath(debug_dir)
    os.makedirs(debug_dir, exist_ok=True)

    def snap(name):
        path = os.path.join(debug_dir, name)
        try:
            driver.save_screenshot(path)
            print(f"  [screenshot] {path}")
        except Exception as e:
            print(f"  [screenshot failed] {e}")

    def click(step_name, settle=1.0):
        """Dismiss any popup/notification, wait for the element, click it.

        Scrolls the element to the middle of the viewport first (not just
        "into view", which can leave it hidden behind a sticky header), and
        falls back to a JS click if a normal click gets intercepted by an
        overlapping element (common on this Vue app's fixed top navbar).
        """
        dismiss_popups(driver)
        kind, val = SELECTORS[step_name]
        el = wait.until(EC.element_to_be_clickable((_by(kind), val)))
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", el
        )
        time.sleep(0.3)
        try:
            el.click()
        except Exception:
            # Fallback: click directly via JS, bypassing whatever element
            # (e.g. the sticky top navbar) intercepted the real click.
            driver.execute_script("arguments[0].click();", el)
        time.sleep(settle)  # let the Vue app re-render before the next lookup
        dismiss_popups(driver)

    # Close any welcome/notice popups that may block the first click
    dismiss_popups(driver)
    time.sleep(1)
    dismiss_popups(driver)

    # 1. Go straight to the report page (confirmed via outerHTML — its <a> has
    # a real href, so no need to click through the ambiguous "Attendance" ->
    # "Attendance" submenu).
    driver.get(BASE_URL + REPORT_PATH)
    time.sleep(2)
    dismiss_popups(driver)
    snap("01_report_page.png")

    # NOTE: there's an unlabeled input field (form-control) on this page whose
    # purpose isn't confirmed yet — this is likely the date range field(s)
    # now targeted by SELECTORS["date_from"]/["date_to"]. Confirm via
    # DevTools and adjust if it turns out to be one combined field instead.
    if from_date or to_date:
        print(f"  Setting date range: {from_date or '(open)'} — {to_date or '(open)'}")
        set_date_range(driver, from_date, to_date)
        snap("01b_after_date_range.png")

    # 2. Click "Search" to load the results (can take up to ~60s for all
    # batches to populate on a slow report)
    print("  Clicking Search — waiting up to 60s for batches/results to load...")
    click("search_button", settle=60.0)
    snap("02_after_search.png")

    # 3. Open the export dropdown
    click("export_dropdown_toggle")
    snap("03_after_export_click.png")

    # 4. Click "Excel" inside the now-visible dropdown
    click("download_button", settle=2.0)
    snap("04_after_excel_click.png")

    time.sleep(3)
    print(f"Download click sequence completed. Debug screenshots in: {debug_dir}")


def wait_for_new_download(before_files, timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        after = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
        new_files = {f for f in (after - before_files) if not f.endswith(".crdownload")}
        if new_files:
            return max(new_files, key=os.path.getmtime)
        time.sleep(1)
    raise TimeoutError("No new .xlsx file appeared in the download folder.")


# ══════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════
def run_downstream_reports(vms_report_path, raw_path, args):
    """After a VMS Report is built, chain the debar list, its notice-board
    copy, and (if a batch list was given) the abstract workbook — used by
    both `format` and `pipeline` so one command produces everything."""
    base, _ = os.path.splitext(os.path.basename(args.output if hasattr(args, "output") else vms_report_path))

    if not args.skip_debar:
        debar_output = args.debar_output or f"{base}_Debar.xlsx"
        dout, log = build_debar_list(vms_report_path, debar_output, args.date)
        print(f"Debar list saved: {dout}")
        for sheet_name, ok, detail in log:
            print(f"  {'OK' if ok else 'SKIPPED'} {sheet_name}: {detail}")

        if not getattr(args, "skip_notice_board", False):
            notice_output = getattr(args, "notice_board_output", None) or f"{base}_Debar_NoticeBoard.xlsx"
            nout, nlog = build_debar_list(vms_report_path, notice_output, args.date, notice_board=True)
            print(f"Notice-board debar list saved: {nout}")
            for sheet_name, ok, detail in nlog:
                print(f"  {'OK' if ok else 'SKIPPED'} {sheet_name}: {detail}")

    if args.batch_list and not args.skip_abstract:
        abstract_output = args.abstract_output or f"{base}_Abstract.xlsx"
        aout, sems = build_abstract(raw_path, args.batch_list, abstract_output, args.date, args.program)
        print(f"Abstract workbook saved (Soft Skill excluded): {aout}")
        for s in sems:
            print(f"  Sheet: SUB {sem_label(s)} SEM")
    elif not args.batch_list:
        print("  (no --batch-list given — skipped the abstract workbook)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="mode", required=True)

    p_format = sub.add_parser("format", help="Format a raw Linways export you already downloaded — also builds the debar list, and the abstract workbook if --batch-list is given")
    p_format.add_argument("input")
    p_format.add_argument("output")
    p_format.add_argument("--low", type=float, default=None, help="Omit to be prompted at runtime")
    p_format.add_argument("--high", type=float, default=None, help="Omit to be prompted at runtime")
    p_format.add_argument("--dept", default="ALL")
    p_format.add_argument("--exclude", default="", help="Comma-separated subjects to drop")
    p_format.add_argument("--include", default="", help="Comma-separated subjects to keep (whitelist, applied before --exclude); default: keep all")
    p_format.add_argument("--batch-list", default=None, help="Path to a BCA/MCA Lab Batch List workbook — cross-references lab/internship faculty & batch by Reg No, and unlocks the abstract workbook")
    p_format.add_argument("--date", default=None, help="'as of' date shown in the debar list and abstract workbook, dd.mm.yyyy (default: today)")
    p_format.add_argument("--program", default="BCA", help="Program name shown in the abstract workbook heading, e.g. 'BCA' or 'MCA'")
    p_format.add_argument("--debar-output", default=None, help="Debar list filename (default: derived from output)")
    p_format.add_argument("--abstract-output", default=None, help="Abstract workbook filename (default: derived from output); only built if --batch-list is given")
    p_format.add_argument("--skip-debar", action="store_true", help="Don't build the debar list")
    p_format.add_argument("--skip-abstract", action="store_true", help="Don't build the abstract workbook even if --batch-list is given")
    p_format.add_argument("--skip-notice-board", action="store_true", help="Don't build the notice-board copy of the debar list")
    p_format.add_argument("--notice-board-output", default=None, help="Notice-board debar list filename (default: derived from --output)")
    ss_format = p_format.add_mutually_exclusive_group()
    ss_format.add_argument("--include-soft-skill", dest="soft_skill", action="store_true", default=None, help="Include Soft Skill in the report")
    ss_format.add_argument("--exclude-soft-skill", dest="soft_skill", action="store_false", help="Drop Soft Skill from the report")
    # Omit both flags to be prompted (y/N) at runtime instead.

    p_download = sub.add_parser("download", help="Only download the raw report from Linways")
    p_download.add_argument("--username", default=os.environ.get("LINWAYS_USERNAME", ""))
    p_download.add_argument("--password", default=os.environ.get("LINWAYS_PASSWORD", ""))
    p_download.add_argument("--show-browser", action="store_true")
    p_download.add_argument("--from-date", default=None, help="Start of date range for the Linways report (format must match the real field once selectors are confirmed)")
    p_download.add_argument("--to-date", default=None, help="End of date range for the Linways report")

    p_debar = sub.add_parser("debar", help="Build the Tentative Debar List from a raw VMS export")
    p_debar.add_argument("input", nargs="?", default=None, help="Omit to auto-use the most recently generated VMS Report")
    p_debar.add_argument("output")
    p_debar.add_argument("--date", default=None, help="'as of' date shown in each title, dd.mm.yyyy (default: today)")

    p_notice = sub.add_parser("notice", help="Build the notice-board copy of the Tentative Debar List (no legend/summary, bigger fonts) from a raw VMS export")
    p_notice.add_argument("input", nargs="?", default=None, help="Omit to auto-use the most recently generated VMS Report")
    p_notice.add_argument("output")
    p_notice.add_argument("--date", default=None, help="'as of' date shown in each title, dd.mm.yyyy (default: today)")
    p_notice.add_argument("--font-scale", type=float, default=1.6, help="Font-size multiplier vs the normal debar list (default 1.6)")
    p_notice.add_argument("--row-scale", type=float, default=1.6, help="Row-height multiplier vs the normal debar list (default 1.6)")

    p_abstract = sub.add_parser("abstract", help="Build the per-semester Abstract workbook (Soft Skill excluded) from a raw consolidated report + Lab Batch List")
    p_abstract.add_argument("input", help="Raw consolidated attendance report")
    p_abstract.add_argument("batch_list", help="BCA/MCA Lab Batch List workbook")
    p_abstract.add_argument("output")
    p_abstract.add_argument("--date", default=None, help="'as of' date shown in the abstract workbook heading, dd.mm.yyyy (default: today)")
    p_abstract.add_argument("--program", default="BCA", help="Program name shown in the abstract workbook heading, e.g. 'BCA' or 'MCA'")

    p_pipeline = sub.add_parser("pipeline", help="Download from Linways, format, build the debar list, and (with --batch-list) the abstract workbook — all in one go")
    p_pipeline.add_argument("--username", default=os.environ.get("LINWAYS_USERNAME", ""))
    p_pipeline.add_argument("--password", default=os.environ.get("LINWAYS_PASSWORD", ""))
    p_pipeline.add_argument("--output", default="VMS_Report.xlsx")
    p_pipeline.add_argument("--low", type=float, default=None, help="Omit to be prompted at runtime")
    p_pipeline.add_argument("--high", type=float, default=None, help="Omit to be prompted at runtime")
    p_pipeline.add_argument("--dept", default="ALL")
    p_pipeline.add_argument("--exclude", default="", help="Comma-separated subjects to drop")
    p_pipeline.add_argument("--include", default="", help="Comma-separated subjects to keep (whitelist, applied before --exclude); default: keep all")
    p_pipeline.add_argument("--batch-list", default=None, help="Path to a BCA/MCA Lab Batch List workbook — cross-references lab/internship faculty & batch by Reg No, and unlocks the abstract workbook")
    p_pipeline.add_argument("--date", default=None, help="'as of' date shown in the debar list and abstract workbook, dd.mm.yyyy (default: today)")
    p_pipeline.add_argument("--program", default="BCA", help="Program name shown in the abstract workbook heading, e.g. 'BCA' or 'MCA'")
    p_pipeline.add_argument("--debar-output", default=None, help="Debar list filename (default: derived from --output)")
    p_pipeline.add_argument("--abstract-output", default=None, help="Abstract workbook filename (default: derived from --output); only built if --batch-list is given")
    p_pipeline.add_argument("--skip-debar", action="store_true", help="Don't build the debar list")
    p_pipeline.add_argument("--skip-abstract", action="store_true", help="Don't build the abstract workbook even if --batch-list is given")
    p_pipeline.add_argument("--skip-notice-board", action="store_true", help="Don't build the notice-board copy of the debar list")
    p_pipeline.add_argument("--notice-board-output", default=None, help="Notice-board debar list filename (default: derived from --output)")
    ss_pipeline = p_pipeline.add_mutually_exclusive_group()
    ss_pipeline.add_argument("--include-soft-skill", dest="soft_skill", action="store_true", default=None, help="Include Soft Skill in the report")
    ss_pipeline.add_argument("--exclude-soft-skill", dest="soft_skill", action="store_false", help="Drop Soft Skill from the report")
    # Omit both flags to be prompted (y/N) at runtime instead.
    p_pipeline.add_argument("--show-browser", action="store_true")
    p_pipeline.add_argument("--from-date", default=None, help="Start of date range for the Linways report (format must match the real field once selectors are confirmed)")
    p_pipeline.add_argument("--to-date", default=None, help="End of date range for the Linways report")

    args = ap.parse_args()

    if args.mode == "format":
        low, high = prompt_for_range(args.low, args.high)
        include_soft_skill = prompt_for_soft_skill(args.soft_skill)
        exclude = [s.strip() for s in args.exclude.split(",") if s.strip()]
        include = [s.strip() for s in args.include.split(",") if s.strip()]
        out, summaries = build_report(args.input, args.output, low, high, args.dept, exclude, include, include_soft_skill, args.batch_list)
        print(f"VMS Report saved: {out}")
        for s in summaries:
            print(f"  {s['Section']}: {s['Count']}")
        run_downstream_reports(out, args.input, args)

    elif args.mode == "debar":
        out, log = build_debar_list(args.input, args.output, args.date)
        print(f"Saved: {out}")
        for sheet_name, ok, detail in log:
            print(f"  {'OK' if ok else 'SKIPPED'} {sheet_name}: {detail}")

    elif args.mode == "notice":
        out, log = build_debar_list(args.input, args.output, args.date, notice_board=True,
                                     font_scale=args.font_scale, row_scale=args.row_scale)
        print(f"Saved: {out}")
        for sheet_name, ok, detail in log:
            print(f"  {'OK' if ok else 'SKIPPED'} {sheet_name}: {detail}")

    elif args.mode == "abstract":
        out, sems = build_abstract(args.input, args.batch_list, args.output, args.date, args.program)
        print(f"Saved: {out}")
        for s in sems:
            print(f"  Sheet: SUB {sem_label(s)} SEM")

    elif args.mode == "download":
        if not args.username or not args.password:
            raise SystemExit("Provide --username/--password or LINWAYS_USERNAME/LINWAYS_PASSWORD env vars.")
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))
        driver = build_driver(headless=not args.show_browser)
        try:
            login(driver, args.username, args.password)
            download_consolidated_report(driver, from_date=args.from_date, to_date=args.to_date)
        finally:
            driver.quit()

        try:
            raw_path = wait_for_new_download(before, timeout=90)
            print(f"SUCCESS — downloaded: {raw_path}")
        except TimeoutError:
            print("NO FILE ARRIVED within 90s. Check the debug_screenshots folder "
                  "next to your downloads folder to see what the page looked like "
                  "at each step.")

    elif args.mode == "pipeline":
        if not args.username or not args.password:
            raise SystemExit("Provide --username/--password or LINWAYS_USERNAME/LINWAYS_PASSWORD env vars.")
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")))

        driver = build_driver(headless=not args.show_browser)
        try:
            login(driver, args.username, args.password)
            download_consolidated_report(driver, from_date=args.from_date, to_date=args.to_date)
        finally:
            driver.quit()

        raw_path = wait_for_new_download(before)
        print(f"Downloaded raw report: {raw_path}")

        low, high = prompt_for_range(args.low, args.high)
        include_soft_skill = prompt_for_soft_skill(args.soft_skill)
        exclude = [s.strip() for s in args.exclude.split(",") if s.strip()]
        include = [s.strip() for s in args.include.split(",") if s.strip()]
        out, summaries = build_report(raw_path, args.output, low, high, args.dept, exclude, include, include_soft_skill, args.batch_list)
        print(f"VMS Report saved: {out}")
        for s in summaries:
            print(f"  {s['Section']}: {s['Count']}")
        run_downstream_reports(out, raw_path, args)


def _running_in_streamlit():
    """True when this file is being executed by `streamlit run` / Streamlit
    Cloud (which loads it as the main module) instead of from a terminal."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__":
    if _running_in_streamlit():
        # Streamlit Cloud was pointed at this file — show the web UI instead
        # of the CLI. Re-executed on every Streamlit rerun.
        import runpy
        runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py"),
                       run_name="__streamlit_app__")
    else:
        main()


# ══════════════════════════════════════════════════════════════════════
# HOW TO FIND THE REAL SELECTORS (needed only for `download`/`pipeline` modes)
# ══════════════════════════════════════════════════════════════════════
# `format` mode above needs none of this — it works right now on any raw
# Linways export you already have.
#
# For `download`/`pipeline`, do this once (~2 minutes):
# 1. Open https://presidencycollege.linways.com/ams/faculty/login in Chrome.
# 2. Right-click the username/email field -> Inspect. Note its id/name.
#    Update SELECTORS["username"] near the top of this file, e.g.:
#      "username": ("ID", "loginId")
# 3. Same for the password field -> SELECTORS["password"].
# 4. Right-click the Login button -> Inspect -> SELECTORS["login_button"].
# 5. After logging in manually, go to wherever you currently download the
#    "Consolidated Course Wise Attendance Report" from (usually a Reports
#    menu). Inspect that nav link -> SELECTORS["reports_nav"] and
#    SELECTORS["consolidated_report_link"].
# 6. Inspect the actual Download/Export button on the report page ->
#    SELECTORS["download_button"].
# 7. For date-range filtering (--from-date/--to-date), inspect the
#    unlabeled form-control field(s) on the report page mentioned in
#    download_consolidated_report() -> SELECTORS["date_from"] and
#    SELECTORS["date_to"]. If it turns out to be ONE combined
#    daterangepicker field rather than two separate inputs, point both
#    entries at that same selector and tell me — set_date_range() will
#    need a small tweak to type "from - to" into it in one go instead of
#    filling two fields.
# 8. Send me the outerHTML (right-click -> Copy -> Copy outerHTML) of each
#    of those elements and I'll fill these in precisely for you.
# ══════════════════════════════════════════════════════════════════════
