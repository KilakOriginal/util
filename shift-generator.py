import sys, os
import pandas as pd
from icalendar import Calendar, Event, vText
import dateutil.parser

def merge_shifts(shifts: list[tuple]) -> dict:
    todo = set([name for (name,_,_,_) in shifts])

    result = {}
    for name in todo:
        result[name] = []

    for shift in shifts:
        name  = shift[0]
        start = shift[1]
        end   = shift[2]
        type  = shift[3]
        
        if result[name] \
        and start == result[name][-1][1] \
        and type == result[name][-1][2]:
            result[name][-1] = (result[name][-1][0], end, type)
        else:
            result[name].append((start, end, type))

    return result

def main():
    if len(sys.argv) < 2:
        sys.exit(f"Usage: {sys.argv[0]} <file_name>")

    path = os.path.abspath(sys.argv[1])
    if not os.path.exists(path):
        sys.exit(f"'{path}' is not a file.")

    try:
        book = pd.ExcelFile(path)
        shift_book = pd.read_excel(book, "Shifts")
    except ValueError:
        sys.exit(f"{path} is not an Excel file.")

    try:
        description_book = pd.read_excel(book, "Descriptions")
    except ValueError:
        print("Warning: 'Shifts' sheet not found.")

    shifts = []
    staff = []
    types = []
    calendar_name = input("Please enter a calendar name: ")
    destination = os.path.abspath(input("Please enter a destination path: "))

    descriptions = {}
    locations = {}

    emails = {}

    try:
        for shift in shift_book["Shift"]:
            types.append(shift)
    except KeyError:
        sys.exit("'Shift' column not found")

    try:
        for persons in shift_book["Staff"]:
            if pd.notnull(persons):
                staff.append([p.strip() for p in persons.split(",")])
            else:
                staff.append([])
    except KeyError:
        sys.exit("'Staff' column not found")

    try:
        for i, time in enumerate(shift_book["Time"]):
            if staff[i]:
                for person in staff[i]: # TODO: handle empty strings
                    k = 1
                    while shift_book["Time"][i + k] == time:
                        k += 1
                    shifts.append((person, time, shift_book["Time"][i + k], types[i]))
    except KeyError:
        sys.exit("'Time' column not found")

    try:
        for (shift, description) in zip(description_book["Shift"],
                                     description_book["Description"]):
            descriptions[shift] = description
    except UnboundLocalError: # Sheet not loaded
        pass
    except KeyError:
        print("Warning: Unable to load shift description.")

    try:
        for (shift, location) in zip(description_book["Shift"],
                                     description_book["Location"]):
            locations[shift] = location
    except UnboundLocalError: # Sheet not loaded
        pass
    except KeyError:
        print("Warning: Unable to load shift location.")

    calendars = merge_shifts(shifts)

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
            event.add("summary", f"{calendar_name}: {entry[2]}")
            try:
                event.add("description", descriptions[entry[2]])
            except KeyError:
                print(f"Skipping description for '{entry[2]}'...")
            event.add("dtstart", dateutil.parser.isoparse(entry[0]))
            event.add("dtend", dateutil.parser.isoparse(entry[1]))
            try:
                event['location'] = vText(locations[entry[2]])
            except KeyError:
                print(f"Skipping location for '{entry[2]}'...")
            calendar.add_component(event)
        
        with open(os.path.join(destination, f"{person}.ics"), "wb") as fs:
            fs.write(calendar.to_ical())

if __name__ == "__main__":
    main()
