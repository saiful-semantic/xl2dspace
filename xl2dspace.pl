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

my $import_dir    = './tmp/import';
my $dspace_dir    = '/opt/dspace';
my $handle_base   = '1';
my $collection_id = '1';
my $eperson_email = 'dspace@localhost';

unless ( $ARGV[0] ) { die "USAGE: $0 <Excel_file>\n\n"; exit; }

my $book;
unless ( $book = ReadData( $ARGV[0] ) ) {
    die "No input Excel file.\n";
}

# First sheet
my $sheet = $book->[1];
warn "total cols: $sheet->{maxcol}\ntotal rows: $sheet->{maxrow}\n";
my @header_row = Spreadsheet::Read::row( $sheet, 1 );

# prepare import dir
unless ( -d $import_dir ) {
    mkdir($import_dir) or die "Couldn't create $import_dir directory, $!";
}
my $coll_dir = $import_dir . '/Col_' . $collection_id;
unless ( -d $coll_dir ) {
    mkdir($coll_dir) or die "Couldn't create $coll_dir directory, $!";
}

# Update upload script
# 	bin/dspace import --add --eperson=joe@user.com --collection=1234/12 --source=Col_12 --mapfile=mapfile
open( UPLOAD, ">>$import_dir/upload.sh" );
print UPLOAD $dspace_dir
  . '/bin/dspace import --add --eperson='
  . $eperson_email;
print UPLOAD ' --collection='
  . $handle_base . '/'
  . $collection_id
  . ' --source=Col_'
  . $collection_id;
print UPLOAD ' --mapfile=mapfile_col_' . $collection_id . "\n";
close(UPLOAD);

# Data rows
my ( $rec_num, $item_num ) = ( 0, 0 );
for my $row_id ( 2 .. $sheet->{maxrow} ) {
    my @row = Spreadsheet::Read::row( $sheet, $row_id );
    $rec_num++;
    $item_num++;

    # Item directory
    my $item_dir = $coll_dir . '/item_' . leftpad_zero( $item_num, 3 );
    unless ( -d $item_dir ) {
        mkdir($item_dir) or die "Couldn't create $item_dir directory, $!";
    }

    my $output = IO::File->new(">$item_dir/dublin_core.xml");
    my $writer =
      XML::Writer->new( OUTPUT => $output, DATA_MODE => 1, DATA_INDENT => 2 );
    $writer->xmlDecl( 'UTF-8', 'no' );
    $writer->startTag( 'dublin_core', "schema" => 'dc' );

    my $file = undef;
  FIELD: for my $col ( 1 .. $sheet->{maxcol} ) {
        next unless $row[$col];
        my $dc_field = $header_row[$col];

        if ( $dc_field eq 'filename' ) {
            $file = $row[$col];
            next FIELD;
        }

        $dc_field =~ s/^dc\.//;
        my $qualifier = 'none';
        ( $dc_field, $qualifier ) = split /\./, $dc_field if $dc_field =~ /\./;

        # Repeatable fields
        if ( $dc_field =~ /^creator$/i ) {
            my @authors = get_array( $row[$col], ';' );
            foreach my $creator (@authors) {
                $writer->dataElement(
                    'dcvalue', $creator,
                    element   => 'creator',
                    qualifier => 'none'
                );
            }
        }
        elsif ( $dc_field =~ /^subject$/i ) {
            my @subjects = get_array( $row[$col], ',' );
            foreach my $keyword (@subjects) {
                $writer->dataElement(
                    'dcvalue', $keyword,
                    element   => 'subject',
                    qualifier => 'none'
                );
            }
        }
        else {
            $writer->dataElement(
                'dcvalue', $row[$col],
                element   => lc($dc_field),
                qualifier => lc($qualifier)
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

sub leftpad_zero {
    my ( $string, $len ) = @_;
    if ( length $string < $len ) {
        return '0' x ( $len - length $string ) . $string;
    }
    return $string;
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
