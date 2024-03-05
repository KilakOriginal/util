## About
Various utility scripts and applications

## Shift Generator
> [!WARNING]
> Outdated

Generate icalendar files from a Microsoft Excel spreadsheet
Feel free to modify is to read csv files instead (Levenshtein distance of 4)

Assumptions:
- The time column is labled "Time"
- The staff column is labled "Staff"
- The shift type column is labled "Shift"
- The last row of each day contains only a time value (last shift ends at)
- Time cells conform to the ISO 8601 standard e.g. 2023-07-16T19:20+01:00
