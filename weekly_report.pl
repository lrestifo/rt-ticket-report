#!/usr/bin/env perl -w
################################################################################
#
# Title:    WEEKLY_REPORT -- Query RT and build an XLSX report out of it
# Author:   Tue Jan 20 21:09:45 CET 2015 lrestifo at esselte dot com
# Description:
#   This script executes the following tasks:
#   1. Reads an SQL query from a text file
#   2. Runs the SQL query against the RT database
#   3. Saves query result set in an XLSX spreadsheet file
#   4. Add formulas to the XLSX spreadheet thus building a complete report
#   5. Saves the resulting XLSX report in a known location
#   This script is designed to be run weekly, and receives the week number as
#   a command-line parameter.  It defines some of the calculation formulas
#   taking the current week into account, and also the resulting spreadsheet
#   output file is named according to the given week
# Usage:
#   weekly_report [ yyyyww ]
#
################################################################################

use strict;
use DBI;
use Excel::Writer::XLSX;

# Database - Customize these
#----------------------------
my $host = "localhost";
my $db   = "rtdb";
my $user = "rtuser";
my $pass = "rtpass";

# File names
#------------
my $sqlf = "./sql/sap_mp_alltickets.sql";
my $xlsx = "./xls/tickets.xlsx";

# Load the RT Query from the SQL source file
#--------------------------------------------
my $sql = "";
my $lin = "";
open( SQL, "<$sqlf" ) or die "Can't open $sqlf: $!\n";
while( $lin = <SQL> ) {
        $sql = $sql . $lin;
}

# Create the Excel workbook
#---------------------------
my $workbook = Excel::Writer::XLSX->new( $xlsx ) or die "Can't open $xlsx: $!\n";
my $datasheet = $workbook->add_worksheet( "RawData" );
my $rowno = 0;
my $colno = 0;
my @row;
my @col;
my $field;

# Connect to RT database and run the query
#------------------------------------------
my $dbh = DBI->connect( "dbi:mysql:".$db.";host=".$host, $user, $pass ) or die "Connection error: " . $DBI::errstr . "\n";
my $sth = $dbh->prepare( $sql )  or die "Prepare statement error: " . $dbh->errstr . "\n";
my $rows = $sth->execute() or die "SQL execution error: " . $sth->errstr . "\n";

# Tell field names apart
#------------------------
foreach $colno( 0 .. $sth->{NUM_OF_FIELDS} ) {
  $col[$colno] = $sth->{NAME}->[$colno];
  $datasheet->write( $rowno, $colno, $col[$colno] );
}

# Let's go
#----------
while( @row = $sth->fetchrow_array() )  {
  $rowno++;
  foreach $field( 0 .. $#row ) {
    $datasheet->write( $rowno, $field, $row[$field] );
  }
}

# Game over
#-----------
$sth->finish;
$workbook->close() or die "Error closing $xlsx: $!\n";
