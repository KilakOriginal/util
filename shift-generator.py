import sys, os
import pandas as pd
from icalendar import Calendar, Event, vText
import dateutil.parser
import operator

def generate_overall_preferences(split_preferences: tuple[tuple, tuple]) -> list:
    preferences = []
    for y in split_preferences[1]:
        for x in split_preferences[0]:
            preferences.append((x,y))
    return preferences

def shifts_unassigned(shifts: dict):
    for shift in shifts:
        if any(shifts[shift]):
            return True
    return False

def slots_free(shifts: dict):
    for shift in shifts:
        if shifts[shift] > 0:
            return True
    return False

def generate_shifts(slots: dict, preferences: dict):
    for key in preferences:
        preferences[key] = generate_overall_preferences(preferences[key])

    to_fill = dict(sorted(slots.items(), key=operator.itemgetter(1), reverse=True))
    shifts = {}

    for slot in to_fill:
        shifts[slot] = []

    while shifts_unassigned(preferences) and slots_free(to_fill):
        for shift in to_fill:
            buffer = []
            for person in preferences:
                if shift in preferences[person]:
                    buffer.append((preferences[person].index(shift), person))
            buffer.sort(key=lambda x: x[0])
            if buffer:
                shifts[shift].append(buffer[0][1])
                to_fill[shift] -= 1
                person = buffer[0][1]
                for i, preference in enumerate(preferences[person]):
                    if preference and preference[1] == shift[1]:
                        preferences[person][i] = None

    return shifts

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

def write_master_table(destination: str, shifts: dict):
    slots = {}
    for shift in shifts:
        (type,(s,e)) = shift
        try:
            slots[type].append((s,e,', '.join(shifts[shift])))
        except KeyError:
            slots[type] = [(s,e,', '.join(shifts[shift]))]

    sheets = []
    for slot in slots:
        slots[slot].sort(key=lambda x: x[0])
        sheets.append((slot, pd.DataFrame([[s,e,n] for (s,e,n) in slots[slot]], columns=["Starts", "Ends", "Staff"])))
    
    with pd.ExcelWriter(destination) as writer:
        for (sheet, df) in sheets:
            df.to_excel(writer, sheet_name=sheet, index=False)

def split_staff_tables(shifts: dict) -> dict:
    result = {}

    for shift in shifts:
        for staff in shifts[shift]:
            (t,(s,e)) = shift
            try:
                result[staff].append((s,e,t))
            except KeyError:
                result[staff] = [(s,e,t)]

    return result

def main():
    s1 = "2024-03-05T08:00+01:00"
    e1 = "2024-03-05T09:00+01:00"
    s2 = "2024-03-05T10:00+01:00"
    e2 = "2024-03-05T11:00+01:00"
    s3 = "2024-03-06T10:00+01:00"
    e3 = "2024-03-06T11:00+01:00"
    
    shifts = generate_shifts({("Bar", (s1, e1)) : 1,
                            ("Bar", (s2, e2)) : 1,
                            ("Einlass", (s1, e1)) : 2},
                            {"Malik" : (("Bar", "Einlass", "Garderobe"),
                                        ((s1, e1), (s2, e2), (s3, e3))),
                            "Bert" : (("Bar", "Einlass", "Garderobe"),
                                      ((s1, e1), (s2, e2), (s3, e3))),
                            "Lara" : (("Einlass", "Bar", "Garderobe"),
                                        ((s2, e2), (s1, e1), (s3, e3))),
                            "Sophie" : (("Garderobe", "Bar", "Einlass"),
                                        ((s2, e2), (s1, e1), (s3, e3)))})

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

    calendar_name = input("Please enter a calendar name: ")
    destination = os.path.abspath(input("Please enter a destination path: "))

    descriptions = {}
    locations = {}

    emails = {}

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

    calendars = split_staff_tables(shifts)

    if not os.path.exists(destination):
        try:
            os.mkdir(destination)
        except Exception as e:
            sys.exit(f"Unable to create directory ({e})")

    write_master_table(os.path.join(destination, "auto.xlsx"), shifts)

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
