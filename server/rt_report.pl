#!/usr/bin/env perl -w
################################################################################
#
# Title:    WEEKLY_REPORT -- Query RT and build an XLSX report out of it
# Author:   Tue Jan 20 21:09:45 CET 2015 lrestifo at esselte dot com
# Description:
#   This script executes the following tasks:
#   1. Reads 2 different SQL queries from 2 text files - summary and detail
#   2. Runs both queries in sequence against the RT database
#   3. Combines result sets from both queries in an XLSX spreadsheet file
#   4. Saves query results as 2 separate JSON documents
#   5. Saves query results as 2 separate CSV files
#   6. Writes the outcome of its task in a JSON object saved as a text file
# Usage:
#   weekly_report [ yyyyww ]
# Notes:
#   * This script is designed to be run weekly without parameters, in which case
#     it computes values based on the current date's week.  In case a numeric
#     parameter is given on the command line, it is interpreted as the year/week
#     number to use for calculations (e.g. 201442 --> Week 42 of year 2014)
#   * CSV output separator is TAB and not comma (in line with RT practice)
# Output:
#   All output is saved in the ./data directory
#   1. tickets_YYYYWW.xls (2 worksheets named Summary and RawData)
#   2. tickets_summary_YYYYWW.json and tickets_rawdata_YYYYWW.json
#   3. tickets_summary_YYYYWW.csv  and tickets_rawdata_YYYYWW.csv
#   4. latest.json
#
################################################################################

use strict;
use Error qw(:try);
use DateTime;
use Getopt::Long;
use RT::Client::REST;
use Excel::Writer::XLSX;

# Check command line parameter
#------------------------------
my $dt = DateTime->now;
my $week = 100 * $dt->week_year() + $dt->week_number();
my $serv = "http://some.server/rt";
my $user = "user";
my $pass = "pass";
GetOptions(
  "week:i"      => \$week,
  "server:s"    => \$serv,
  "user:s"      => \$user,
  "password:s"  => \$pass
) or die("Error in command line arguments\n");

my $rt = RT::Client::REST->new(
  server => $serv,
  timeout => 30
);

try {
  $rt->login(username => $user, password => $pass);
} catch Exception::Class::Base with {
  die "Can't login as '$user': ", shift->message;
};

try {
  # Get ticket #10
  my $ticket = $rt->show(type => 'ticket', id => 42310);
  print "Id: ", $ticket->{id}, "\n";
  print "Subject: ", $ticket->{Subject}, "\n";
  print "Owner: ", $ticket->{Owner}, "\n";
  print "Queue: ", $ticket->{Queue}, "\n";
  print "Requestors: ", $ticket->{Requestors}, "\n";
  print "Priority: ", $ticket->{Priority}, "\n";
  print "Status: ", $ticket->{Status}, "\n";
  print "Created: ", $ticket->{Created}, "\n";
  print "Started: ", $ticket->{Started}, "\n";
  print "Due: ", $ticket->{Due}, "\n";
  print "LastUpdated: ", $ticket->{LastUpdated}, "\n";
  print "Resolved: ", $ticket->{Resolved}, "\n";
  print "Country: ", $ticket->{"CF.{Country}"}, "\n";
  print "Request Type: ", $ticket->{"CF.{Request_Type}"}, "\n";
  print "Impact: ", $ticket->{"CF.{Impact Scope}"}, "\n";
  print "Classification: ", $ticket->{"CF.{Ticket Classification}"}, "\n";
} catch RT::Client::REST::UnauthorizedActionException with {
  print "You are not authorized to view ticket #10\n";
} catch RT::Client::REST::Exception with {
  # something went wrong.
  print shift->message;
};

# Get all Swiss tickets in reverse order:
my @ids = $rt->search(
  type => "ticket",
  query => "(Status = 'new' OR Status = 'open' OR Status = 'stalled' OR Status = 'user-testing') AND (Requestor.Organization = 'Esselte Switzerland')",
  orderby => "-id"
);
for my $id (@ids) {
  my $t = $rt->show(type => "ticket", id => $id);
  print "#$id: ", $t->{Subject}, "\n";
}

# Database - Customize these
#----------------------------
# my $host = "localhost";
# my $port = 3306;
# my $db   = "rtdb";
# my $user = "rtuser";
# my $pass = "rtpass";

# File names
#------------
# my $sqld = "./sql/sap_mp_alltickets.sql";
# my $sqls = "./sql/sap_mp_summary.sql";
# my $xlsx = "./data/tickets_$yyww.xlsx";
# my $json1 = "./data/tickets_summary_$yyww.json";
# my $json2 = "./data/tickets_rawdata_$yyww.json";
# my $csv1 = "./data/tickets_summary_$yyww.csv";
# my $csv2 = "./data/tickets_rawdata_$yyww.csv";
# my $conf = "./data/latest.json";
# 
# # Read the 2 queries from the SQL source files
# #----------------------------------------------
# my $sqd = "";
# my $sqs = "";
# # Detail
# my $lin = "";
# open( SQL, "<$sqld" ) or die "Can't open $sqld: $!\n";
# while( $lin = <SQL> ) {
#   $sqd = $sqd . $lin;
# }
# # Summary - hacking the year/week as necessary
# my $now = "YEARWEEK(NOW(), 3)";
# my $txt = "";
# open( SQL, "<$sqls" ) or die "Can't open $sqls: $!\n";
# while( $txt = <SQL> ) {
#   $txt =~ s/\Q$now\E/$week/g unless $week == 0;
#   $sqs = $sqs . $txt;
# }
# 
# # Create the Excel workbook
# #---------------------------
# my $workbook = Excel::Writer::XLSX->new( $xlsx ) or die "Can't open $xlsx: $!\n";
# my $summary = $workbook->add_worksheet( "Summary" );
# my $rawdata = $workbook->add_worksheet( "RawData" );
# my $rowno = 0;
# my $colno = 0;
# my @row;
# my @col;
# my @csvcol;
# my $csvrow;
# 
# # Connect to RT database
# #------------------------
# my $dbh = DBI->connect( "dbi:mysql:".$db.";host=".$host, $user, $pass ) or die "Connection error: " . $DBI::errstr . "\n";
# 
# #
# # Populate the Summary worksheet and csv
# #----------------------------------------
# my $sths = $dbh->prepare( $sqs )  or die "Prepare statement error [summary]: " . $dbh->errstr . "\n";
# my $summ = $sths->execute() or die "SQL execution error [summary]: " . $sths->errstr . "\n";
# open( CSV1, ">$csv1" ) or die "Can't open $csv1: $!\n";
# # Column titles
# $rowno = 0;
# foreach $colno( 0 .. $sths->{NUM_OF_FIELDS}-1 ) {
#   $col[$colno] = $sths->{NAME}->[$colno];
#   $col[$colno] =~ s/_/ /g;
#   $summary->write( $rowno, $colno, $col[$colno] );
# }
# @csvcol = @{$sths->{NAME}};
# $csvrow = join( "\t", @csvcol );
# print CSV1 $csvrow . "\n";
# # Data rows
# while( @row = $sths->fetchrow_array() )  {
#   $rowno++;
#   foreach $colno( 0 .. $#row ) {
#     $summary->write( $rowno, $colno, $row[$colno] );
#   }
#   $csvrow = join( "\t", @row );
#   print CSV1 $csvrow . "\n";
# }
# my $summaryRows = $rowno;
# # Attributes
# $summary->set_tab_color( "green" );
# $summary->set_column(  0,  0, 30 ); # Department
# $summary->set_column(  1, 12, 12 ); # Value area
# $summary->set_column( 13, 13, 15 ); # YYYYWW
# $summary->set_row( 0, 30 );         # Column titles
# # Done with this
# close( CSV1 ) or die "Error closing $csv1: $!\n";
# $sths->finish;
# 
# #
# # Populate the RawData worksheet and csv
# #----------------------------------------
# my $sthd = $dbh->prepare( $sqd )  or die "Prepare statement error [detail]: " . $dbh->errstr . "\n";
# my $rows = $sthd->execute() or die "SQL execution error [detail]: " . $sthd->errstr . "\n";
# open( CSV2, ">$csv2" ) or die "Can't open $csv2: $!\n";
# # Column titles
# $rowno = 0;
# foreach $colno( 0 .. $sthd->{NUM_OF_FIELDS}-1 ) {
#   $col[$colno] = $sthd->{NAME}->[$colno];
#   $rawdata->write( $rowno, $colno, $col[$colno] );
# }
# @csvcol = @{$sthd->{NAME}};
# $csvrow = join( "\t", @csvcol );
# print CSV2 $csvrow . "\n";
# # Data rows
# $rowno = 0;
# while( @row = $sthd->fetchrow_array() )  {
#   $rowno++;
#   foreach $colno( 0 .. $#row ) {
#     $rawdata->write( $rowno, $colno, $row[$colno] );
#   }
#   my $csvrow = join( "\t", @row );
#   print CSV2 $csvrow . "\n";
# }
# my $rawdataRows = $rowno;
# # Attributes
# $rawdata->set_tab_color( "black" );
# # Done with this
# close( CSV2 ) or die "Error closing $csv2: $!\n";
# $sthd->finish;
# 
# #
# # Workbook attributes
# #---------------------
# $workbook->set_properties(
#   title     => "SAP Ticket Report Week $yyww",
#   author    => "Luciano Restifo",
#   company   => "Esselte Leitz GmbH & Co KG",
#   subject   => "Report specification: Esselte CIO",
#   comments  => "v0.1 - Feb 2015"
# );
# $workbook->close() or die "Error closing $xlsx: $!\n";
# 
# #
# # Now repeat the queries and save results as JSON data
# #------------------------------------------------------
# open( JSON1, ">$json1" ) or die "Can't open $json1: $!\n";
# open( JSON2, ">$json2" ) or die "Can't open $json2: $!\n";
# my $dsn = "dbname=$db;host=$host;port=$port";
# print JSON1 DBIx::JSON->new( $dsn, "mysql", $user, $pass )->do_select( $sqs )->get_json;
# print JSON2 DBIx::JSON->new( $dsn, "mysql", $user, $pass )->do_select( $sqd )->get_json;
# close( JSON1 ) or die "Error closing $json1: $!\n";
# close( JSON2 ) or die "Error closing $json2: $!\n";
# 
# #
# # Save run results in the configuration file
# #--------------------------------------------
# open( CONF, ">$conf" ) or die "Can't open $conf: $!\n";
# print CONF "{ " .
#     "\"week\":$yyww, " .
#     "\"summaryRows\":$summaryRows, \"rawdataRows\":$rawdataRows, " .
#     "\"xlsx\":\"$xlsx\", " .
#     "\"summaryJson\":\"$json1\", \"rawdataJson\":\"$json2\", " .
#     "\"summaryCsv\":\"$csv1\", \"rawdataCsv\":\"$csv2\" " .
#   "}\n";
# close( CONF ) or die "Error closing $conf: $!\n";

# Game over
#-----------
exit( 0 );
