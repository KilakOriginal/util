'''
Assumptions:
- The leftmost column contains time cells
- All other columns contain the names of people
- The First row does not contain any relevant information
    - It may be used for labels or left blank
- The last row of each day contains only a time value (last shift ends at)
- Time cells conform to the ISO 8601 standard e.g. 2023-07-16T19:20+01:00
'''

import sys, os
import pandas as pd
from icalendar import Calendar, Event, vText
import dateutil.parser

def main():
    if len(sys.argv) < 2:
        sys.exit(f"Usage: {sys.argv[0]} <file_name>")

    path = os.path.abspath(sys.argv[1])
    if not os.path.exists(path):
        sys.exit(f"'{path}' is not a file.")

    try:
        shifts = pd.read_excel(path)
    except ValueError:
        sys.exit(f"{path} is not an Excel file.")

    times = [dateutil.parser.isoparse(shifts.iloc[i,0]) 
             for i in range(len(shifts))]
    persons = []
    for i in range(1, shifts.shape[1]):
        persons.append([shifts.iloc[j,i] for j in range(len(shifts))])
    calendar_name = input("Please enter a calendar name: ")
    event_name = input("Please enter an event name: ")
    description = input("Please enter a description: ")
    location = input("Please enter the event location: ")
    destination = os.path.abspath(input("Please enter the destination path: "))

    calendars = {} # name:[time]

    for position in persons:
        current_person = ''
        for i, person in enumerate(position):
            if pd.notnull(person):
                if person == current_person:
                    try:
                        calendars[person][-1] = (calendars[person][-1][0],
                                                 times[i+1])
                    except KeyError:
                        calendars[person] = [(times[i], times[i+1])]
                else:
                    try:
                        calendars[person].append((times[i], times[i+1]))
                    except KeyError:
                        calendars[person] = [(times[i], times[i+1])]
                current_person = person

    if not os.path.exists(destination):
        try:
            os.mkdir(destination)
        except Exception as e:
            sys.exit(f"Unable to create directory ({e})")

    for person in calendars:
        calendar = Calendar()

        calendar.add("prodid", "-//shift-generator//https://wwww.github.com/KilakOriginal///")
        calendar.add("version", "2.0")
        calendar.add("name", calendar_name)

        for entry in calendars[person]:
            event = Event()
            event.add("summary", event_name)
            event.add("description", description)
            event.add("dtstart", entry[0])
            event.add("dtend", entry[1])
            event['location'] = vText(location)
            calendar.add_component(event)
        
        with open(os.path.join(destination, f"{person}.ics"), "wb") as fs:
            fs.write(calendar.to_ical())

if __name__ == "__main__":
    main()
