# Timesheet Validator with Google Calendar

## Overview
This Flask app validates employee timesheets against Google Calendar events 
and generates an Excel report of missing and extra entries.

## Features
- Upload CSV timesheet.
- Fetch Google Calendar events.
- Compare timesheet vs calendar.
- Download Excel validation report.

## How to Run

1. Clone the project or unzip the folder.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt

Run the app:

python app.py

Open in browser: http://127.0.0.1:5000