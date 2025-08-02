import csv
from datetime import datetime, timezone, timedelta
from google_calender import get_calendar_events  # Make sure this function exists


def parse_timesheet(file_path):
    entries = []
    ist = timezone(timedelta(hours=5, minutes=30))  # India Standard Time
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                start_time = datetime.strptime(f"{row['date']} {row['start']}", "%Y-%m-%d %H:%M").replace(tzinfo=ist)
                end_time = datetime.strptime(f"{row['date']} {row['end']}", "%Y-%m-%d %H:%M").replace(tzinfo=ist)

                entries.append({
                    "date": row['date'],
                    "start": start_time,
                    "end": end_time,
                    "project": row['project']
                })
            except Exception as e:
                print(f"Skipping invalid row: {row}, error: {e}")
    return entries

def validate_timesheet(timesheet_file):
    timesheet_entries = parse_timesheet(timesheet_file)
    calendar_events = get_calendar_events()

    missing_entries = []
    extra_entries = []

    for entry in timesheet_entries:
        matched = False
        for event in calendar_events:
            if entry['start'] >= event['start'] and entry['end'] <= event['end']:
                matched = True
                break
        if not matched:
            extra_entries.append(entry)

    for event in calendar_events:
        matched = False
        for entry in timesheet_entries:
            if entry['start'] >= event['start'] and entry['end'] <= event['end']:
                matched = True
                break
        if not matched:
            missing_entries.append(event)

    return {"missingEntries": missing_entries, "extraEntries": extra_entries}

if __name__ == "__main__":
    csv_file = "sample_timesheet.csv"
    print(f"Validating timesheet: {csv_file} ...")

    report = validate_timesheet(csv_file)

    if not report["missingEntries"] and not report["extraEntries"]:
        print("✅ No discrepancies found! Timesheet matches calendar.")
    else:
        print("\n=== Validation Report ===")
        if report["missingEntries"]:
            print("\nMissing Entries (in calendar but not in timesheet):")
            for e in report["missingEntries"]:
                print(f"  - {e['start']} to {e['end']} : {e.get('summary','No Title')}")
        else:
            print("\nNo missing entries.")

        if report["extraEntries"]:
            print("\nExtra Entries (in timesheet but not in calendar):")
            for e in report["extraEntries"]:
                print(f"  - {e['start']} to {e['end']} : {e['project']}")
        else:
            print("\nNo extra entries.")
