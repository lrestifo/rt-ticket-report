# rt-ticket-report
Run SQL against Best Practical's Request Tracker database and save results as MSExcel

This project consists of a script that executes the following tasks:

1. Reads 2 different SQL queries from 2 text files - summary and detail
2. Runs both queries against the RT database in sequence
3. Combines result sets from both queries in an XLSX spreadsheet file
4. Add formulas to the XLSX spreadsheet thus building a complete report
5. Saves the resulting XLSX report in a known location

This script is designed to be run weekly, and calculates its values based on the current date's week.
In case an optional parameter is given, it is interpreted as an Year/Week number to use for the calculation.

## Usage:
weekly_report [ yyyyww ]

## Issues:
When the script runs, Perl may complain about locale settings.
Look [here](https://sskaje.me/2014/01/lc-ctype-issue/) for a discussion of the topic and how to turn locale settings off.

## Credits:
Uses the great [Excel::Writer::XLSX](http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX-0.81/) Perl module.
