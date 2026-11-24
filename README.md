# Employee Data Excel App

## Run
1. Extract the ZIP.
2. Open `Employee_Excel_Data_App`.
3. Double-click `run.bat`.
4. On first run it creates `.venv` and installs dependencies.

## Manual run
```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

## Fields
- Employee Name
- Employee ID
- Mobile No
- Aadhaar No
- Address
- Email ID

Data is saved to `data/employees.xlsx`.
