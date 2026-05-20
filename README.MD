# E-Samarth Marks Entry Assistant
The e-samarth web portal marks entry assistant. Teachers have excel sheet of seat number and marks which needs to be entered on the portal.

## The Flow
- The Teacher opens this app and loads the excel sheet with seat number and marks.
- The the Teachers clicks a button in the app to open chrome browser with debugging 
- The teachers logs in and reaches the page where the data entry is to be done.
- The teacher clicks the button to select the table where data entry is to be performed
- The teacher selects the column where seat number is present
- Then clicks the Start button on the app to starts data entry
- Progress bar show the status
- The app has a tabular list of data imported from the excel and an additional column with the status of the data entry like PENDING, DONE, NOT FOUND

## Sample Data
Please check the `sample_data` folder for sample forms and screenshot of the table. The third column is the `Exam Roll Number` which is the Seat Number of the student. The fifth column is where the marks are to be entered and a check box is to be pressed and then save button is to be pressed for every student.

## App Details
Build a Python GUI app with Playwright.

## Implemented Stack
- GUI: PyQt6
- Automation: Playwright (Chromium)
- Data import: pandas + openpyxl

## Run The App
1. Install dependencies:
	- `pip install -r requirements.txt`
2. Optional: install Playwright browser (only needed if installed Google Chrome is unavailable):
	- `playwright install chromium`
3. Start the app:
	- `python app.py`

## Build One-File EXE Manually
1. Create and activate `.venv` if not already done.
2. Install app dependencies:
	- `pip install -r requirements.txt`
3. Install PyInstaller:
	- `pip install pyinstaller`
4. Add an icon file named `app.ico` in the project root.
5. Run this command from the project root:

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --name "E-Samarth-Marks-Entry-Assistant" --icon "icon.ico" --hidden-import playwright.sync_api  app.py
```

Generated executable:
- `dist\E-Samarth-Marks-Entry-Assistant.exe`

Notes:
- The packaged app still prefers installed Google Chrome on the target machine.
- If Chrome is not installed on that machine, install Playwright Chromium separately or update the app to use a bundled browser.
- The one-file build starts slower on first launch because PyInstaller extracts bundled files to a temporary directory.

## Usage
1. Click `Browse` and select your Excel/CSV.
2. Click `Load Data`, then choose seat number and marks columns.
3. Click `Open Chromium (Debug)` and login to the e-samarth portal.
4. Navigate to the marks entry page.
5. Click `Select Table`, then click the marks table in the portal.
6. Click `Select Seat Column`, then click a seat number cell in that table.
7. Keep `Marks Column Index in Portal` as `5` unless your page layout differs.
8. Click `Start Entry`.

The app updates each row with `PENDING`, `DONE`, `NOT FOUND`, or `ERROR` and shows progress.