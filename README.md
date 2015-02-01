# rt-ticket-report
Run SQL against Best Practical's Request Tracker database and save results as MSExcel

This project consists of a script that executes the following tasks:

1. Reads 2 different SQL queries from 2 text files - summary and detail
2. Runs both queries in sequence against the RT database
3. Combines result sets from both queries in an XLSX spreadsheet file
4. Saves query results as 2 separate JSON documents
5. Saves query results as 2 separate CSV files
6. Writes the outcome of its work in a JSON configuration object saved also as a text file

This script is designed to be run weekly, and calculates its values based on the current date's week.
In case an optional parameter is given, it is interpreted as an Year/Week number to use for the calculation.

## Usage:
weekly_report [ yyyyww ]

## Output:
All output is saved in the ./data directory.  Files names as follows:

1. tickets_YYYYWW.xls (2 worksheets named Summary and RawData)
2. tickets_summary_YYYYWW.json + tickets_rawdata_YYYYWW.json
3. tickets_summary_YYYYWW.csv  + tickets_rawdata_YYYYWW.csv
4. latest.json

*NOTE*: CSV fields are separated by a TAB character (not comma).
This is in line with RT practice.

## Issues:
When the script runs, Perl may complain about locale settings.
Look [here](https://sskaje.me/2014/01/lc-ctype-issue/) for a discussion of the topic and how to turn locale settings off.

## Credits:
Uses the great [Excel::Writer::XLSX](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.81/) Perl module.
Uses also DBIx::JSON.
