#!/usr/bin/perl
# Name: xl2dspace.pl
# Author: Saiful Amin <saiful@semanticconsulting.com>
# Description: Converts MS Excel file into DSpace Simple Archive Format

use strict;
use warnings;
use Spreadsheet::Read;
use XML::Writer;
use IO::File;
use File::Copy qw(copy);
use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
use Archive::Zip::Tree;

my $import_dir    = './import';
my $dspace_dir    = '/opt/dspace';
my $handle_base   = '1';
my $collection_id = '2';
my $eperson_email = 'dspace@localhost';

unless ( $ARGV[0] ) { die "USAGE: $0 <Spreadsheet_file>\n\n"; exit; }
my $book  = ReadData( $ARGV[0] );
my $sheet = $book->[1];             # First sheet
warn "total cols: $sheet->{maxcol}\ntotal rows: $sheet->{maxrow}\n";

# prepare import dir
unless ( -d $import_dir ) {
    mkdir($import_dir) or die "Couldn't create $import_dir directory, $!\n";
}
my $coll_dir = $import_dir . '/Col_' . $collection_id;
unless ( -d $coll_dir ) {
    mkdir($coll_dir) or die "Couldn't create $coll_dir directory, $!\n";
}

# Create upload script
#  bin/dspace import --add --eperson=joe@user.com --collection=1234/12 --source=Col_12 --mapfile=mapfile
open( UPLOAD, ">$import_dir/upload.sh" );
print UPLOAD $dspace_dir
  . '/bin/dspace import --add --eperson='
  . $eperson_email
  . ' --collection='
  . $handle_base . '/'
  . $collection_id
  . ' --source=Col_'
  . $collection_id
  . ' --mapfile=mapfile_col_'
  . $collection_id . "\n";
close(UPLOAD);

# Data rows
my @header_row = Spreadsheet::Read::row( $sheet, 1 );
my ( $rec_num, $item_num ) = ( 0, 0 );
for my $row_id ( 2 .. $sheet->{maxrow} ) {
    my @row = Spreadsheet::Read::row( $sheet, $row_id );
    $rec_num++;
    $item_num++;

    # Item directory
    my $item_dir = $coll_dir . '/item_' . sprintf("%03d", $item_num);
    unless ( -d $item_dir ) {
        mkdir($item_dir) or die "Couldn't create $item_dir directory, $!";
    }

    my $output = IO::File->new(">$item_dir/dublin_core.xml");
    my $writer =
      XML::Writer->new( OUTPUT => $output, DATA_MODE => 1, DATA_INDENT => 2 );
    $writer->xmlDecl( 'UTF-8', 'no' );
    $writer->startTag( 'dublin_core', "schema" => 'dc' );

    my $file = undef;
  FIELD: for my $col ( 0 .. $sheet->{maxcol} ) {
        next unless $row[$col];
        my ($dc_field, $qualifier, $delimiter) = parse_header($header_row[$col]);
        next unless $dc_field;

        if ( $dc_field eq 'filename' ) {
            $file = $row[$col];
            next FIELD;
        }
        
        # Repeatable fields
        if ( $delimiter ) {
            my @all_values = get_array( $row[$col], $delimiter );
            foreach my $value (@all_values) {
                $value =~ s/^\s*//; # cleanup, just in case.
                $writer->dataElement(
                    'dcvalue', $value,
                    element   => $dc_field,
                    qualifier => $qualifier
                );
            }
        }
        # Non-repeatable ones
        else {
            $writer->dataElement(
                'dcvalue', $row[$col],
                element   => $dc_field,
                qualifier => $qualifier
            );
        }

    }

    $writer->endTag("dublin_core");
    $writer->end();
    $output->close();

    # files & manifest
    open( MANIFEST, ">$item_dir/contents" );

    # Full-text content
    my $file_loc = 'repository/' . $file if $file;
    if ( $file && -e $file_loc ) {
        print MANIFEST $file . "\tbundle:ORIGINAL\n";
        copy $file_loc, $item_dir;
    }

    # default license
    copy 'license.txt', "$item_dir/license.txt";
    print MANIFEST "license.txt\tbundle:LICENSE";
    close(MANIFEST);

    warn "$rec_num records\n" if ( $rec_num % 100 == 0 );
}

my $zip = Archive::Zip->new();
$zip->addTree( $coll_dir, 'collection_1' );
unless ( $zip->writeToFileNamed( $import_dir . '/archive.zip') == AZ_OK ) {
	die "Error creating archive.zip file: $!\n";
}

# convert string to array by delimiter
sub get_array {
    my ( $str, $delimiter ) = @_;
    my @array;
    if ( $str =~ m/$delimiter/ ) {
        @array = split "$delimiter", $str;
    }
    else {
        push @array, $str;
    }
    return @array;
}

sub parse_header {
    my $field = shift;
    $field =~ s/^dc\.//;
    
    my $delimit = '';
    if ($field =~ /^([^\(]+)\s*\((.+)\)$/) {
        ($field, $delimit) = ($1, $2);
    }
    
    my $qualifier = 'none';
    ($field, $qualifier ) = split /\./, $field if $field =~ /\./;
    return (lc($field), lc($qualifier), $delimit);
}
