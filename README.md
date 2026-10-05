# Linways Attendance -> VMS Report

## Files
- `vms_pipeline.py`      main script (format, debar, notice, abstract, download, pipeline)
- `presidency_logo.jpg`  put YOUR logo here, next to the script (optional; sheets are built without it if missing)
- `requirements.txt`     Python dependencies

## Setup
    pip install -r requirements.txt

Selenium modes also need Google Chrome installed.

## Credentials (download / pipeline only)
    # Windows (PowerShell)
    $env:LINWAYS_USERNAME="you@presidency.edu.in"
    $env:LINWAYS_PASSWORD="your-password"
    # macOS / Linux
    export LINWAYS_USERNAME="you@presidency.edu.in"
    export LINWAYS_PASSWORD="your-password"

## Common commands
    # Format an already-downloaded export (also builds debar list + notice board copy)
    python vms_pipeline.py format raw.xlsx VMS_Report.xlsx --low 0 --high 75

    # Same, plus the Abstract workbook (Soft Skill is always excluded from it)
    python vms_pipeline.py format raw.xlsx VMS_Report.xlsx --batch-list "Lab Batch List.xlsx" --date 07.09.2026

    # Download from Linways and do everything
    python vms_pipeline.py pipeline --output VMS_Report.xlsx --from-date 01/06/2026 --to-date 31/07/2026

    # Individual builders
    python vms_pipeline.py debar  VMS_Report.xlsx Debar.xlsx
    python vms_pipeline.py notice VMS_Report.xlsx Notice.xlsx
    python vms_pipeline.py abstract raw.xlsx "Lab Batch List.xlsx" Abstract.xlsx --program BCA

## Output folders (created next to where you run the script)
- `VMS Report/`
- `Tentative Debar List/`
- `Tentative Debar List (Notice Board)/`
- `Abstract Report/`
- `downloads/` and `debug_screenshots/` (Selenium modes)
