import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime
import os
import argparse

def csv_to_ics(input_file, output_file=None):
    """
    Converts a CSV file containing work schedules into an ICS file for calendar import.
    """

    try:
        # Load the CSV file
        schedule_df = pd.read_csv(input_file)

        # Check required columns exist
        required_columns = {"First Name", "Last Name", "Shift Region", "Start Date", "Start Time", "End Date", "End Time", "Hours"}
        if not required_columns.issubset(schedule_df.columns):
            missing_cols = required_columns - set(schedule_df.columns)
            raise ValueError(f"Missing columns in CSV: {missing_cols}")

        # Create a new calendar
        cal = Calendar()

        # Function to convert a row of data into an iCalendar event
        def add_event_to_calendar(row):
            try:
                event = Event()
                event.add('summary', f"{row['First Name']} {row['Last Name']} - {row['Shift Region']}")

                # Parse date and time properly
                start_dt = datetime.strptime(f"{row['Start Date']} {row['Start Time']}", '%Y-%m-%d %I:%M:%S %p')
                end_dt = datetime.strptime(f"{row['End Date']} {row['End Time']}", '%Y-%m-%d %I:%M:%S %p')

                event.add('dtstart', start_dt)
                event.add('dtend', end_dt)
                event.add('description', f"Hours: {row['Hours']}")
                event.add('location', row['Shift Region'])
                
                cal.add_component(event)

            except Exception as e:
                print(f"Error processing row {row.to_dict()}: {e}")

        # Iterate over each row in the CSV and add to the calendar
        schedule_df.apply(add_event_to_calendar, axis=1)

        # Generate output file name if not provided
        if output_file is None:
            output_file = os.path.splitext(input_file)[0] + '.ics'

        # Save the ICS file
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())

        print(f"✅ ICS file successfully created: {output_file}")

    except FileNotFoundError:
        print(f"❌ Error: Input file '{input_file}' not found.")
    except Exception as e:
        print(f"❌ An error occurred: {str(e)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert CSV schedule to ICS file for Google Calendar import.")
    parser.add_argument("input_file", help="Path to the input CSV file")
    parser.add_argument("-o", "--output_file", help="Path to save the output ICS file (optional)")
    
    args = parser.parse_args()
    csv_to_ics(args.input_file, args.output_file)
