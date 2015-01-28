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
# Usage:
#   weekly_report [ yyyyww ]
#
#   This script is designed to be run weekly without parameters, in which case
#   it computes values based on the current date's week.  In case a numeric
#   parameter is given on the command line, it is interpreted as the year/week
#   number to use for calculations (e.g. 201442 --> Week 42 of year 2014)
#
################################################################################

use strict;
use DBI;
use Excel::Writer::XLSX;

# Check command line parameter
#------------------------------
my $week = ( $#ARGV == 0 && $ARGV[0] =~ /^\d{6}$/ ? $ARGV[0] : 0 );

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

# Read the 2 queries from the SQL source files
#----------------------------------------------
my $sqd = "";
my $sqs = "";
# Detail
my $lin = "";
open( SQL, "<$sqld" ) or die "Can't open $sqld: $!\n";
while( $lin = <SQL> ) {
  $sqd = $sqd . $lin;
}
# Summary - hacking the year/week as necessary
my $now = "YEARWEEK(NOW(), 3)";
my $txt = "";
open( SQL, "<$sqls" ) or die "Can't open $sqls: $!\n";
while( $txt = <SQL> ) {
  $txt =~ s/\Q$now\E/$week/g unless $week == 0;
  $sqs = $sqs . $txt;
}

# Create the Excel workbook
#---------------------------
my $workbook = Excel::Writer::XLSX->new( $xlsx ) or die "Can't open $xlsx: $!\n";
my $summary = $workbook->add_worksheet( "Summary" );
my $rawdata = $workbook->add_worksheet( "RawData" );
my $rowno = 0;
my $colno = 0;
my @row;
my @col;

# Connect to RT database
#------------------------
my $dbh = DBI->connect( "dbi:mysql:".$db.";host=".$host, $user, $pass ) or die "Connection error: " . $DBI::errstr . "\n";

#
# Populate the Summary worksheet
#--------------------------------
my $sths = $dbh->prepare( $sqs )  or die "Prepare statement error [summary]: " . $dbh->errstr . "\n";
my $summ = $sths->execute() or die "SQL execution error [summary]: " . $sths->errstr . "\n";
# Column titles
$rowno = 0;
foreach $colno( 0 .. $sths->{NUM_OF_FIELDS}-1 ) {
  $col[$colno] = $sths->{NAME}->[$colno];
  $col[$colno] =~ s/_/ /g;
  $summary->write( $rowno, $colno, $col[$colno] );
}
# Data rows
$rowno = 0;
while( @row = $sths->fetchrow_array() )  {
  $rowno++;
  foreach $colno( 0 .. $#row ) {
    $summary->write( $rowno, $colno, $row[$colno] );
  }
}
# Attributes
$summary->set_tab_color( 'green' );

#
# Populate the RawData worksheet
#--------------------------------
my $sthd = $dbh->prepare( $sqd )  or die "Prepare statement error [detail]: " . $dbh->errstr . "\n";
my $rows = $sthd->execute() or die "SQL execution error [detail]: " . $sthd->errstr . "\n";
# Column titles
$rowno = 0;
foreach $colno( 0 .. $sthd->{NUM_OF_FIELDS}-1 ) {
  $col[$colno] = $sthd->{NAME}->[$colno];
  $rawdata->write( $rowno, $colno, $col[$colno] );
}
# Data rows
$rowno = 0;
while( @row = $sthd->fetchrow_array() )  {
  $rowno++;
  foreach $colno( 0 .. $#row ) {
    $rawdata->write( $rowno, $colno, $row[$colno] );
  }
}
# Attributes
$summary->set_tab_color( 'black' );

# Game over
#-----------
$sthd->finish;
$sths->finish;
$workbook->close() or die "Error closing $xlsx: $!\n";
