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
#   4. Saves query results as JSON documents
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
use DateTime;
use DBI;
use DBIx::JSON;
use Excel::Writer::XLSX;

# Check command line parameter
#------------------------------
my $dt = DateTime->now;
my $week = ( $#ARGV == 0 && $ARGV[0] =~ /^\d{6}$/ ? $ARGV[0] : 0 );
my $yyww = ( $week == 0 ? 100 * $dt->week_year() + $dt->week_number() : $week );

# Database - Customize these
#----------------------------
my $host = "localhost";
my $port = 3306;
my $db   = "rtdb";
my $user = "rtuser";
my $pass = "rtpass";

# File names
#------------
my $sqld = "./sql/sap_mp_alltickets.sql";
my $sqls = "./sql/sap_mp_summary.sql";
my $xlsx = "./data/tickets_$yyww.xlsx";
my $conf = "./data/config.json";
my $json1 = "./data/tickets_summary_$yyww.json";
my $json2 = "./data/tickets_rawdata_$yyww.json";

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
# Populate the Summary worksheet / table
#----------------------------------------
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
while( @row = $sths->fetchrow_array() )  {
  $rowno++;
  foreach $colno( 0 .. $#row ) {
    $summary->write( $rowno, $colno, $row[$colno] );
  }
}
# Attributes
$summary->set_tab_color( "green" );
$summary->set_column(  0,  0, 30 ); # Department
$summary->set_column(  1, 12,  5 ); # Value area
$summary->set_column( 13, 13,  8 ); # YYYYWW
$summary->set_row( 0, 30 );         # Column titles

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
$rawdata->set_tab_color( "black" );

#
# Workbook attributes
#---------------------
$workbook->set_properties(
  title     => "SAP Ticket Report Week $yyww",
  author    => "Luciano Restifo",
  company   => "Esselte Leitz GmbH & Co KG",
  comments  => "v0.1 - Feb 2015"
);

#
# Now repeat the queries and save results as JSON data
#------------------------------------------------------
open( JSON1, ">$json1" ) or die "Can't open $json1: $!\n";
open( JSON2, ">$json2" ) or die "Can't open $json2: $!\n";
my $dsn = "dbname=$db;host=$host;port=$port";
print JSON1 DBIx::JSON->new( $dsn, "mysql", $user, $pass )->do_select( $sqs )->get_json;
print JSON2 DBIx::JSON->new( $dsn, "mysql", $user, $pass )->do_select( $sqd )->get_json;
# Save the week number in the configuration file
open( CONF, ">$conf" ) or die "Can't open $conf: $!\n";
print CONF "{ \"week\":$yyww, \"summary\":\"$json1\", \"rawdata\":\"$json2\" }\n";
close( CONF ) or die "Error closing $conf: $!\n";

# Game over
#-----------
$sthd->finish;
$sths->finish;
$workbook->close() or die "Error closing $xlsx: $!\n";
close( JSON1 ) or die "Error closing $json1: $!\n";
close( JSON2 ) or die "Error closing $json2: $!\n";
