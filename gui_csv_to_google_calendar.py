import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime
import os
import pickle

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Define Google Calendar API scopes
SCOPES = ["https://www.googleapis.com/auth/calendar"]

# Authenticate and get Google Calendar service
def authenticate_google_calendar():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)

# Convert CSV to Google Calendar events
def csv_to_google_calendar():
    input_file = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not input_file:
        return

    service = authenticate_google_calendar()

    try:
        schedule_df = pd.read_csv(input_file)

        required_columns = {"First Name", "Last Name", "Shift Region", "Start Date", "Start Time", "End Date", "End Time", "Hours"}
        if not required_columns.issubset(schedule_df.columns):
            missing_cols = required_columns - set(schedule_df.columns)
            messagebox.showerror("Error", f"Missing columns in CSV: {missing_cols}")
            return

        # Iterate over each row and add to Google Calendar
        for _, row in schedule_df.iterrows():
            try:
                start_dt = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", '%Y-%m-%d %I:%M:%S %p')
                end_dt = datetime.strptime(f"{row['End Date']} {row['End Time']}", '%Y-%m-%d %I:%M:%S %p')

                event = {
                    'summary': f"{row['First Name']} {row['Last Name']} - {row['Shift Region']}",
                    'location': row['Shift Region'],
                    'description': f"Hours: {row['Hours']}",
                    'start': {
                        'dateTime': start_dt.isoformat(),
                        'timeZone': 'America/New_York',  # Adjust to your timezone
                    },
                    'end': {
                        'dateTime': end_dt.isoformat(),
                        'timeZone': 'America/New_York',
                    },
                }

                event_result = service.events().insert(calendarId='primary', body=event).execute()
                print(f"Event created: {event_result.get('htmlLink')}")

            except Exception as e:
                print(f"Error processing row {row.to_dict()}: {e}")

        messagebox.showinfo("Success", "All events have been added to Google Calendar!")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI Setup
root = tk.Tk()
root.title("CSV to Google Calendar")

btn = tk.Button(root, text="Select CSV & Upload to Google Calendar", command=csv_to_google_calendar, padx=20, pady=10)
btn.pack(pady=20)

root.mainloop()