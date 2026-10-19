# Control Tower – Warehouse Operations Dashboard (Streamlit)

## Quick start (Windows)
1. Install Python 3.10+ from python.org (tick "Add Python to PATH").
2. Unzip this folder, e.g. to C:\warehouse_dashboard
3. Copy WAREHOUSE_Visibility_rename.xlsx into the same folder (optional – you can upload it in the app instead).
4. Double-click run_dashboard.bat. First run installs packages (2–5 min).
5. The browser opens at http://localhost:8501

## Quick start (Mac / Linux)
    cd warehouse_dashboard
    ./run_dashboard.sh

## Manual commands
    python -m venv .venv
    .venv\Scripts\activate          (Mac/Linux: source .venv/bin/activate)
    pip install -r requirements.txt
    streamlit run app.py

## Updating data
Paste new exports into the input sheets of the workbook (GRN-Input, PutawayReturn-Input,
Picking-Input, Packed-Input), update plan/manpower rows in Main Table, save, then refresh
the browser (or re-upload). Plan/manpower can also be edited in the "Plan & manpower" tab.

## Files
- app.py – dashboard UI
- engine.py – calculations (same logic as the Excel Formula Guide)
- requirements.txt – Python packages
- .streamlit/config.toml – theme + 500 MB upload limit
