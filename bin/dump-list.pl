#!/usr/bin/perl -I/var/wwwsites/SharePerltopus/perl-lib


use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;

use Getopt::Long;

Getopt::Long::Configure("permute");

my $show_help=0;
my $show_version=0;
my $debug=0;

GetOptions( 
	'help|?' => \$show_help,
	'version' => \$show_version,
	'debug' => \$debug,
) || show_usage();


my( $site, $list_id, $view_id ) = @ARGV;

my $spc = new SharePoint::Connection(
	debug=>$debug,
	"site" => $site,
	"credentials" => "etc/totl.credentials",
);


my $map = {};
my $listinfo =  $spc->GetList( $list_id );
foreach my $field ( @{$listinfo->{fields}} )
{
	$map->{"ows_".$field->{StaticName}} = $field->{DisplayName};
}
my @items =  $spc->GetListItems( $list_id, $view_id );

foreach my $item ( @items )
{
	print "\n";
	my $cells = {};
	CELL: foreach my $k ( sort keys %{$item} )
	{
		next CELL unless defined $map->{$k};
		my $v = $item->{$k};
		my $name = $map->{$k};
		$v =~ s/^.*;#//;
		$cells->{$name} = sprintf( "%30s : %s\n",$name,$v );
	}
	foreach my $cell_name ( sort keys %$cells )
	{
		print $cells->{$cell_name};
	}
	
}

exit;

