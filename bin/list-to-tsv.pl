#!/usr/bin/perl -I/var/wwwsites/SharePerltopus/perl-lib


use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;
use SharePoint::CommandLine;

my %opts = SharePoint::CommandLine->get_options( 
	{ fields=>"" }, 
	{ "fields-file:s" => "fields" } );
my( $list_id, $view_id ) = @ARGV;

binmode( STDOUT, ":utf8" );

my @fields = ();
open( F, $opts{fields} ) || die "can't read '$opts{fields}': $!";
while( my $line = <F> )
{
	chomp $line;
	push @fields, $line;
}
close F;

my $spc = new SharePoint::Connection( %opts );

my $map = {};
my $listinfo =  $spc->GetList( $list_id );
foreach my $field ( @{$listinfo->{fields}} )
{
	$map->{"ows_".$field->{StaticName}} = $field->{DisplayName};
}
my @items =  $spc->GetListItems( $list_id, $view_id );

print join( "\t", @fields )."\n";
foreach my $item ( @items )
{
	my $cells = {};
	CELL: foreach my $k ( sort keys %{$item} )
	{
		next CELL unless defined $map->{$k};
		my $v = $item->{$k};
		my $name = $map->{$k};
		$v =~ s/^.*;#//;
		$cells->{$name} = $v;
	}
	my @values = ();
	foreach my $cell_name ( @fields )
	{
		my $v = $cells->{$cell_name};
		$v = "" unless defined $v;
		push @values, $v;
	}
	print join( "\t", @values )."\n";
}

exit;

