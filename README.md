# Timesheet Validator with Google Calendar
This project is a Timesheet Validation System built with Python Flask (Backend) and HTML/CSS (Frontend).
It allows users to upload their timesheet CSV file, validates the records against Google Calendar events, and generates an Excel report for missing and extra entries.

#Features
Backend: Python Flask

Frontend: HTML & CSS (directly served by Flask)

Timesheet CSV Upload

Google Calendar Integration

Validation of Timesheet Entries

Excel Report Generation (Missing & Extra Entries)

Directly integrated frontend and backend (single Flask app)

#Step-by-Step Project Workflow
1. Upload Timesheet
User uploads a CSV file from the frontend.

CSV Format Example:

date	start	end	project
2025-07-01	09:00	11:00	Project A

2. Fetch Google Calendar Events
Backend fetches events using the google_calender.py script.

Events are normalized to Indian Standard Time (IST).

3. Validate Timesheet
Compare each timesheet entry with calendar events.

Identify:

Extra Entries: Entries not in Google Calendar

Missing Entries: Calendar events not logged in timesheet

4. Generate Excel Report
Backend generates validation_report.xlsx with:

Extra Entries Sheet

Missing Entries Sheet

Summary if no discrepancies found

5. Download Report
After validation, user can download the Excel report directly from the browser.


Tech Stack
Python 3

Flask (Backend Framework)

HTML, CSS (Frontend)

Pandas (Excel report generation)

OpenPyXL (Excel writing)

How to Run
Install dependencies:
pip install flask pandas openpyxl google-api-python-client

Run the Flask server:
python app.py

Open in browser:
http://127.0.0.1:5000

Output
Upload CSV → Validate → Download Excel report.

Example report sheets:

Extra Entries

Missing Entries
