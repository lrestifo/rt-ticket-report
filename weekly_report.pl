#!/usr/bin/env perl -w
################################################################################
#
# Title:    WEEKLY_REPORT -- Query RT and build an XLSX report out of it
# Author:   Tue Jan 20 21:09:45 CET 2015 lrestifo at esselte dot com
# Description:
#   This script executes the following tasks:
#   1. Reads 2 different SQL queries from 2 text files - summary and detail
#   2. Runs both queries against the RT database in sequence
#   3. Combines result sets from both queries in an XLSX spreadsheet file
#   4. Add formulas to the XLSX spreadheet thus building a complete report
#   5. Saves the resulting XLSX report in a known location
#   This script is designed to be run weekly, and calculates its values
#   based on the current date's week.  The logic is defined in the SQL
# Usage:
#   weekly_report
#
################################################################################

use strict;
use DBI;
use Excel::Writer::XLSX;

# Check command line parameters
#-------------------------------
# die "Usage: weekly_report YYYYWW\n E.g.: weekly_report 201503 generates report for week 3 of year 2015\n" unless $#ARGV == 0 && $ARGV[0] =~ /^\d{6}$/;
# my $week = $ARGV[0];

# Database - Customize these
#----------------------------
my $host = "localhost";
my $db   = "rtdb";
my $user = "rtuser";
my $pass = "rtpass";

# File names
#------------
my $sqld = "./sql/sap_mp_alltickets.sql";
my $sqls = "./sql/sap_mp_summary.sql";
my $xlsx = "./xls/tickets.xlsx";

# Load the 2 queries from the SQL source files
#----------------------------------------------
my $sqd = "";
my $sqs = "";
my $lin = "";
open( SQL, "<$sqld" ) or die "Can't open $sqld: $!\n";
while( $lin = <SQL> ) {
  $sqd = $sqd . $lin;
}
$lin = "";
open( SQL, "<$sqls" ) or die "Can't open $sqls: $!\n";
while( $lin = <SQL> ) {
  $sqs = $sqs . $lin;
}

# Create the Excel workbook
#---------------------------
my $workbook = Excel::Writer::XLSX->new( $xlsx ) or die "Can't open $xlsx: $!\n";
my $datasheet = $workbook->add_worksheet( "RawData" );
my $summary = $workbook->add_worksheet( "Summary" );
my $rowno = 0;
my $colno = 0;
my @row;
my @col;

# Connect to RT database and run the queries
#--------------------------------------------
my $dbh = DBI->connect( "dbi:mysql:".$db.";host=".$host, $user, $pass ) or die "Connection error: " . $DBI::errstr . "\n";
my $sthd = $dbh->prepare( $sqd )  or die "Prepare statement error [detail]: " . $dbh->errstr . "\n";
my $rows = $sthd->execute() or die "SQL execution error [detail]: " . $sthd->errstr . "\n";
my $sths = $dbh->prepare( $sqs )  or die "Prepare statement error [summary]: " . $dbh->errstr . "\n";
my $summ = $sths->execute() or die "SQL execution error [summary]: " . $sths->errstr . "\n";


# Use field names for the column titles
#---------------------------------------
$rowno = 0;
foreach $colno( 0 .. $sthd->{NUM_OF_FIELDS} ) {
  $col[$colno] = $sthd->{NAME}->[$colno];
  $datasheet->write( $rowno, $colno, $col[$colno] );
}
$rowno = 0;
foreach $colno( 0 .. $sths->{NUM_OF_FIELDS} ) {
  $col[$colno] = $sths->{NAME}->[$colno];
  $summary->write( $rowno, $colno, $col[$colno] );
}

# Let's go
#----------
$rowno = 0;
while( @row = $sthd->fetchrow_array() )  {
  $rowno++;
  foreach $colno( 0 .. $#row ) {
    $datasheet->write( $rowno, $colno, $row[$colno] );
  }
}
$rowno = 0;
while( @row = $sths->fetchrow_array() )  {
  $rowno++;
  foreach $colno( 0 .. $#row ) {
    $summary->write( $rowno, $colno, $row[$colno] );
  }
}

# Game over
#-----------
$sthd->finish;
$sths->finish;
$workbook->close() or die "Error closing $xlsx: $!\n";
