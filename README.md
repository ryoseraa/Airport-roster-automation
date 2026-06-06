# Airport Roster Automation — CCSI Airport, Lucknow

Automated weekly duty roster generator for the Air Traffic Systems 
(CNS) Department, Chaudhary Charan Singh International Airport, Lucknow.

Replaces a manual Excel-based scheduling process for 50+ staff across 6 units.

## How to run
1. `pip install openpyxl`
2. Place your 4 input Excel files in the same folder as the script
3. `python generate_roster.py`
4. Enter the week start date when prompted (format: YYYY-MM-DD)

## Input files required
- `SHIFT.xlsx` — shift duty employees
- `GEN-EMP.xlsx` — general duty employees  
- `leave.xlsx` — leave records with start/end dates
- `ADDITIONALDUTY.xlsx` — additional duty / overtime records

## Shift cycle logic
B (07:00–13:00) → C (13:00–19:00) → DA (19:00–07:00) → OFF → repeat

> Note: All input files are excluded from this repo. Demo data only.
