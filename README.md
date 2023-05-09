# Peer Review Combinator

Canvas's anonymous peer review would make the one being reviewed anonymous. This makes sense
if you use it to do assignment peer grading - the students don't know who they are grading so
they can't just give their friends good grades. However, for our senior design teamwork
assessment, we want to make the reviewers anonymous so they can give their honest opnions
about their teammates. The way we do it is to ask the students to submit their assessments
in an Excel spreadsheet. This application would combine these spreadsheets into one, which
can then be imported into [Ascent](https://ascent.cysun.org).

## Configuration

Copy `appsettings.json.sample` to `appsettings.json` and change the following settings:

- `RosterFile` - The class roster. This can be downloaded from Canvas using the Export
  Grade Book function. The exported file is actually a CSV file - it needs to converted
  into a proper Excel file with a couple of rows removed ("points possible" and "Test
  Student"). We need this file to get the campus ids of the students.
- `InputFolder` - This is where the downloaded spreadsheets are.
- `OutpuFile` - The combined spreadsheet.
- `ExpectedColumns` - The expected columns in the spreadsheets.
- `OptionalColumnCount` - The number of columns that are optional (e.g. Comments).
  This is used in the code to determine whether a row should be considered empty.
