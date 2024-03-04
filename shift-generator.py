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
        if shifts[shift]:
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
                del_list = []
                for preference in preferences[person]:
                    if preference[1] == shift[1]:
                        del_list.append(preference)
                for element in del_list:
                    preferences[person].remove(element)

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

def main():
    '''
    shifts = generate_shifts({("Shift A", ("Start a", "End a")) : 1,
                            ("Shift A", ("Start b", "End b")) : 1,
                            ("Shift B", ("Start a", "End a")) : 2},
                            {"Malik" : (("Shift A", "Shift B", "Shift C"),
                                        (("Start a", "End a"), ("Start b", "End b"), ("Start c", "End c"))),
                            "Bert" : (("Shift A", "Shift B", "Shift C"),
                                      (("Start a", "End a"), ("Start b", "End b"), ("Start c", "End c"))),
                            "Lara" : (("Shift A", "Shift B", "Shift C"),
                                        (("Start b", "End b"), ("Start a", "End a"), ("Start c", "End c"))),
                            "Sophie" : (("Shift A", "Shift B", "Shift C"),
                                        (("Start b", "End b"), ("Start a", "End a"), ("Start c", "End c")))})
    
    print(shifts)
    '''
    
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
