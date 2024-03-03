## About
Various utility scripts and applications

## Shift Generator
Generate icalendar files from a Microsoft Excel spreadsheet
Feel free to modify is to read csv files instead (Levenshtein distance of 4)

Assumptions:
- The leftmost column contains time cells
- All other columns contain the names of people
- The First row does not contain any relevant information
    - It may be used for labels or left blank
- The last row of each day contains only a time value (last shift ends at)
- Time cells conform to the ISO 8601 standard e.g. 2023-07-16T19:20+01:00