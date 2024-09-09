import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime

# Load your CSV file
schedule_df = pd.read_csv('schedule-2024-09-01--2024-12-31--Muhammad Usman-Khan.csv')

# Create a new calendar
cal = Calendar()

# Function to convert a row of data into an iCalendar event
def add_event_to_calendar(row):
    event = Event()
    event.add('summary', f"{row['First Name']} {row['Last Name']} - {row['Shift Region']}")
    event.add('dtstart', datetime.strptime(f"{row['Start Date']} {row['Start Time']}", '%Y-%m-%d %I:%M:%S %p'))
    event.add('dtend', datetime.strptime(f"{row['End Date']} {row['End Time']}", '%Y-%m-%d %I:%M:%S %p'))
    event.add('description', f"Hours: {row['Hours']}")
    event.add('location', row['Shift Region'])
    cal.add_component(event)

# Iterate over the dataframe rows and add events to the calendar
for index, row in schedule_df.iterrows():
    add_event_to_calendar(row)

# Save the calendar to an .ics file
with open('schedule-2024-09-01--2024-12-31--Muhammad Usman-Khan.ics', 'wb') as f:
    f.write(cal.to_ical())

print("ICS file created successfully.")