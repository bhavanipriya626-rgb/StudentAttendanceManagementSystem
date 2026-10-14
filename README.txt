# Student Attendance Management System

## Features
- Add and list students
- Mark daily attendance by class and date
- Save student and attendance data in `attendance_data.xlsx`
- Update attendance for a date/class without creating duplicate rows
- Monthly student-wise attendance percentage
- Download monthly summary and detailed Excel reports

## Requirements
- Python 3.10 or newer recommended
- Internet connection for the initial package installation

## Windows setup
1. Extract this project ZIP to a folder, for example `Student_Attendance_Project`.
2. Open that folder in VS Code, or open PowerShell in that folder.
3. Install packages:

   ```powershell
   py -m pip install -r requirements.txt
   ```

4. Run the app:

   ```powershell
   py -m streamlit run app.py
   ```

5. Streamlit will display a local URL, usually `http://localhost:8501`. Open it in your browser.

## How to use
1. Open **Students** and add each student. Roll numbers must be unique.
2. Open **Mark Attendance**, choose the date and class, set each student's status, then click **Save Attendance**.
3. Open **Monthly Reports**, choose a month and class, review the percentage, and download an Excel report.

## Excel storage
The app creates `attendance_data.xlsx` in the same folder where you run the app. It contains:
- `Students` sheet
- `Attendance` sheet

Attendance percentage = present days / days with saved attendance records × 100.
Only dates for which attendance was saved are included. Save attendance for every relevant class day for the percentage to be meaningful.

## Important
- Do not open and edit `attendance_data.xlsx` in Excel while the app is saving data.
- Back up the workbook regularly.
- The workbook is stored locally on this computer; it is not automatically synchronized online.
