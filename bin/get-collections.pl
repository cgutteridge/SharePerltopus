#!/usr/bin/perl -I/var/wwwsites/SharePerl/perl-lib


use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;

my( $site ) = @ARGV;

require "secrets.pl";
my $spc = new SharePoint::Connection(
	debug=>0,
	"site" => $site,
	get_credentials(),
);

my @result = $spc->GetListCollection();
print "** List Collections **\n";
foreach my $data (@result) 
{
	print $data->{Title}." - ".$data->{ID}."\n";
}
print "\n";



exit;

