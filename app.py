from flask import Flask, request, render_template, send_file, url_for
import csv
from datetime import datetime, timezone, timedelta
from google_calender import get_calendar_events
import pandas as pd
import os

# ------------------ Config ------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

REPORT_FILE = os.path.join(BASE_DIR, "validation_report.xlsx")
IST = timezone(timedelta(hours=5, minutes=30))  # Indian Standard Time

# Initialize Flask with templates folder
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))


# ------------------ Helper Functions ------------------

def parse_timesheet(file_path):
    """Parse the uploaded CSV file and convert times to IST."""
    entries = []
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                start_time = datetime.strptime(
                    f"{row['date']} {row['start']}", "%Y-%m-%d %H:%M"
                ).replace(tzinfo=IST)
                end_time = datetime.strptime(
                    f"{row['date']} {row['end']}", "%Y-%m-%d %H:%M"
                ).replace(tzinfo=IST)

                entries.append({
                    "date": row['date'],
                    "start": start_time,
                    "end": end_time,
                    "project": row.get('project', 'N/A')
                })
            except Exception as e:
                print(f"Skipping invalid row: {row}, error: {e}")
    return entries


def normalize_calendar_events(events):
    """Convert Google Calendar events to IST and datetime objects."""
    normalized_events = []
    for event in events:
        try:
            start = event['start']
            end = event['end']

            if isinstance(start, str):
                start = datetime.fromisoformat(start.replace("Z", "+00:00"))
            if isinstance(end, str):
                end = datetime.fromisoformat(end.replace("Z", "+00:00"))

            normalized_events.append({
                "start": start.astimezone(IST),
                "end": end.astimezone(IST),
                "summary": event.get("summary", "No Title")
            })
        except Exception as e:
            print(f"Skipping invalid calendar event: {event}, error: {e}")
    return normalized_events


def validate_timesheet(timesheet_file):
    """Compare timesheet entries against Google Calendar events."""
    timesheet_entries = parse_timesheet(timesheet_file)
    calendar_events = get_calendar_events()

    if not calendar_events:
        return {"error": "No calendar events fetched."}

    calendar_events = normalize_calendar_events(calendar_events)

    missing_entries, extra_entries = [], []

    # Identify extra timesheet entries
    for entry in timesheet_entries:
        matched = any(
            entry['start'] >= event['start'] and entry['end'] <= event['end']
            for event in calendar_events
        )
        if not matched:
            extra_entries.append(entry)

    # Identify missing entries
    for event in calendar_events:
        matched = any(
            entry['start'] >= event['start'] and entry['end'] <= event['end']
            for entry in timesheet_entries
        )
        if not matched:
            missing_entries.append(event)

    return {"missingEntries": missing_entries, "extraEntries": extra_entries}


def create_excel_report(report_data):
    """Generate Excel report for missing and extra entries."""
    def serialize_list(data_list):
        return [
            {k: (v.isoformat() if isinstance(v, datetime) else v) for k, v in item.items()}
            for item in data_list
        ]

    extra = serialize_list(report_data.get("extraEntries", []))
    missing = serialize_list(report_data.get("missingEntries", []))

    df_extra = pd.DataFrame(extra)
    df_missing = pd.DataFrame(missing)

    with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as writer:
        sheets_written = False

        if not df_extra.empty:
            df_extra.to_excel(writer, sheet_name="Extra Entries", index=False)
            sheets_written = True
        if not df_missing.empty:
            df_missing.to_excel(writer, sheet_name="Missing Entries", index=False)
            sheets_written = True

        # ✅ Ensure at least one visible sheet exists
        if not sheets_written:
            pd.DataFrame([["No discrepancies found"]]).to_excel(
                writer, sheet_name="Summary", index=False, header=False
            )


# ------------------ Routes ------------------

@app.route("/", methods=["GET"])
def home():
    return render_template("frontend.html")


@app.route("/validate-timesheet", methods=["POST"])
def upload_timesheet():
    if "file" not in request.files:
        return render_template("frontend.html", message="No file uploaded")

    file = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    report = validate_timesheet(file_path)
    create_excel_report(report)

    return render_template(
        "frontend.html",
        message="Validation completed! Download report below.",
        download_url=url_for("download_report")
    )


@app.route("/download-report", methods=["GET"])
def download_report():
    if not os.path.exists(REPORT_FILE):
        return "No report generated yet. Please validate a timesheet first.", 404
    return send_file(REPORT_FILE, as_attachment=True, download_name="validation_report.xlsx")


# ------------------ Main Entry ------------------

if __name__ == "__main__":
    app.run(debug=True, port=5000)
