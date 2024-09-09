import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime
import argparse
import os

def csv_to_ics(input_file, output_file=None):
    # Load CSV file
    schedule_df = pd.read_csv(input_file)

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

    # Generate output file name if not provided
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + '.ics'

    # Save the calendar to an .ics file
    with open(output_file, 'wb') as f:
        f.write(cal.to_ical())

    print(f"ICS file created successfully: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert CSV schedule to ICS file.")
    parser.add_argument("input_file", help="Input CSV file path")
    parser.add_argument("-o", "--output_file", help="Output ICS file path (optional)")
    args = parser.parse_args()

    try:
        csv_to_ics(args.input_file, args.output_file)
    except FileNotFoundError:
        print(f"Error: Input file '{args.input_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")