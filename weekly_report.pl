#!/usr/bin/env perl -w
################################################################################
#
# Title:
# Author:
# Description:
#
################################################################################
#
use strict;
use DBI;
use Excel::Writer::XLSX;

# DATABASE - CUSTOMISE THESE
#----------------------------
my $host = 'localhost';
my $db   = 'rtdb';
my $user = 'rtuser';
my $pass = 'rtpass';

# Load the RT Query from the SQL source file
#--------------------------------------------
my $sql = "";
my $lin = "";
open(SQL, "<./sql/sap_mp_alltickets.sql") or die "Can't open ./sql/sap_mp_alltickets.sql: $!\n";
while($lin = <SQL>) {
        $sql = $sql . $lin;
}

# connect to RT database
#------------------------
my $dbh = DBI->connect('dbi:mysql:'.$db.';host='.$host, $user, $pass) or die "Connection Error: $DBI::errstr\n";

# connect to RT database
#------------------------
my $dbh = DBI->connect('dbi:mysql:'.$db.';host='.$host, $user, $pass) or die "Connection Error: $DBI::errstr\n";

# run the query and save results as Excel
#-----------------------------------------
my $ss = Spreadsheet::WriteExcel::FromDB->read($dbh, $sql);
$ss->write_xls('./data/rtdata.xls');

# now reexecute, retrieve, output data
#--------------------------------------
my $sth = $dbh->prepare($sql) or die "Prepare Statement Error:" . $dbh->errstr . "\n";
my $rows = $sth->execute() or die "SQL Execution Error:" . $sth->errstr . "\n";

# Tell field names apart
#------------------------
foreach $colno (0 .. $sth->{NUM_OF_FIELDS}) {
        $col[$colno] = $sth->{NAME}->[$colno];
}

# Let's go
#----------
open RTDATA, ">./data/rtdata.json" or die "Can't write ./data/rtdata.json: $!\n";
open RTMONGO, ">./data/rtdata.mongoimport.json" or die "Can't write ./data/rtdata.mongoimport.json: $!\n";
print RTDATA "[";
while (@row = $sth->fetchrow_array())  {
        # Output the entire row on both RTDATA and RTMONGO
        print  RTDATA '{ ';
        print RTMONGO '{ ';
        foreach $field (0 .. $#row) {
                print  RTDATA '"' . $col[$field] . '":';
                print RTMONGO '"' . $col[$field] . '":';
                if ( $ffmt{$col[$field]} =~ /\$date/ ) {
                        print  RTDATA sprintf "new Date(%lu)", $row[$field] * 1000;
                        print RTMONGO sprintf $ffmt{$col[$field]}, $row[$field] * 1000;
                } else {
                        if ( $ffmt{$col[$field]} =~ /\$subject/ ) {
                                print  RTDATA $json->encode($row[$field]);
                                print RTMONGO $json->encode($row[$field]);
                        } else {
                                print  RTDATA sprintf $ffmt{$col[$field]}, $row[$field];
                                print RTMONGO sprintf $ffmt{$col[$field]}, $row[$field];
                        }
                }
                if ($field < $#row) {
                        print  RTDATA ', ';
                        print RTMONGO ', ';
                }
        }
        print RTMONGO " }\n";
        print  RTDATA " }\n";
        $created = DateTime->from_epoch( epoch => $row[$IX_CREA] );
        $yy_index = $created->year();                                                           # e.g. 2012
        $mm_index = 100 * $created->year() + $created->month();         # e.g. 201209
        $rgn_index = $row[$IX_REGN];            # Region indicator
        $ctg_index = $row[$IX_CATE];            # Problem category
        $tkt_yy{$yy_index}{"Europe"}++;
        $tkt_mm{$mm_index}{"Europe"}++;
        $tkt_yy{$yy_index}{$rgn_index}++;
        $tkt_mm{$mm_index}{$rgn_index}++;
        $tkt_categ{$yy_index}{"Europe"}{$ctg_index}++;
        $tkt_categ{$yy_index}{$rgn_index}{$ctg_index}++;
}
$sth->finish;

# ------


# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( 'perl.xlsx' );

# Add a worksheet
$worksheet = $workbook->add_worksheet();

#  Add and define a format
$format = $workbook->add_format();
$format->set_bold();
$format->set_color( 'red' );
$format->set_align( 'center' );

# Write a formatted and unformatted string, row and column notation.
$col = $row = 0;
$worksheet->write( $row, $col, 'Hi Excel!', $format );
$worksheet->write( 1, $col, 'Hi Excel!' );

# Write a number and a formula using A1 notation
$worksheet->write( 'A3', 1.2345 );
$worksheet->write( 'A4', '=SIN(PI()/4)' );