#!/usr/bin/perl -w
############################################################################
#    Excel2CSV Converter: convert XLS and XLSX (OpenXML) to CSV            #
#    Copyright (C) 2014-2015 by Natale Vinto aka bluesman                  #
#    ebballon@gmail.com                                                    #
#                                                                          #
#    This program is free software; you can redistribute it and#or modify  #
#    it under the terms of the GNU General Public License as published by  #
#    the Free Software Foundation; either version 2 of the License, or     #
#    (at your option) any later version.                                   #
#                                                                          #
#    This program is distributed in the hope that it will be useful,       #
#    but WITHOUT ANY WARRANTY; without even the implied warranty of        #
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         #
#    GNU General Public License for more details.                          #
#                                                                          #
#    You should have received a copy of the GNU General Public License     #
#    along with this program; if not, write to the                         #
#    Free Software Foundation, Inc.,                                       #
#    59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.             #
############################################################################

# v. 0.1 Convert only the first sheet  

use strict;
use warnings;

use Text::Iconv;
use Spreadsheet::XLSX;

my $xls = '';
my $csv = '';
my $DEBUG = 0;
my $VERSION = "0.1";

sub help {
	print "Excel2CSV $VERSION by blues-man\nUsage: excel2csv.pl File.xlsx\n";
	exit
}

if (scalar(@ARGV) != 1){
	help();
} else {
	$xls = $ARGV[0];
}
my @tmp = split('\.', $xls);
help() unless $tmp[1] =~ /xls/;

$csv = $tmp[0].".csv";

$SIG{'__WARN__'} = sub { warn $_[0] unless (caller eq "Spreadsheet::XLSX"); };


open(CSV, ">$csv") or die("Unable to create file $csv $@");
my $converter = Text::Iconv -> new ("utf-8", "windows-1251");



my $excel = Spreadsheet::XLSX -> new ($xls, $converter);
print "I will work only on the first sheet, sorry.\n" if $excel->worksheets() > 1;
 
my $sheet = $excel->worksheet(0);
 
printf("Sheet: %s\n", $sheet->{Name}) if $DEBUG;
$sheet -> {MaxRow} ||= $sheet -> {MinRow};
my $dcsv = '';

 foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
		my $line = '';
 
		$sheet -> {MaxCol} ||= $sheet -> {MinCol};
		
		foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
				my $cell = $sheet -> {Cells} [$row] [$col];                        
				if ($cell) {
					$line.=$cell-> {Val}.";";
					printf("( %s , %s ) => %s\n", $row, $col, $cell -> {Val}) if $DEBUG;
				} else {
					$line.=";";
				}

		}
		
		$line =~ s/;$//g;
		$line.="\n";
		print CSV $line;
		$dcsv.=$line;
		

}
        
print $dcsv if $DEBUG;
close(CSV);
print "CSV created: $csv\n";

