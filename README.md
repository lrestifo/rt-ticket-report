# rt-ticket-report
Run SQL against Best Practical's Request Tracker database and save results as MSExcel

This project consists of a script that executes the following tasks:

1. Reads an SQL query from a text file
2. Runs the SQL query against the RT database
3. Saves query result set in an XLSX spreadsheet file
4. Add formulas to the XLSX spreadsheet thus building a complete report
5. Saves the resulting XLSX report in a known location

The script is designed to be run weekly, and receives the week number as a command-line parameter.
It defines some of the calculation formulas by taking the week number into account, and also the resulting spreadsheet output file is named according to the given week.

## Usage:
weekly_report [ yyyyww ]

## Issues:
When the script runs, Perl may complain about locale settings.
Look [here](https://sskaje.me/2014/01/lc-ctype-issue/) for a discussion of the topic and how to turn locale settings off.
