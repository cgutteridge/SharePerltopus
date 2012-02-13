#!/usr/bin/perl -I/var/wwwsites/SharePerl/perl-lib

use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;

require "secrets.pl";
my $spc = sp_connect();
my $info =  $spc->GetList( "{58F8FB8E-04E3-4FC3-9ADB-DEE1B04BC346}" );
exit;

my @result = $spc->GetListCollection();
print "** List Collections **\n";
foreach my $data (@result) 
{
	print $data->{Title}." - ".$data->{ID}."\n";
}
print "\n";


my @items = $spc->GetListItems( "{58F8FB8E-04E3-4FC3-9ADB-DEE1B04BC346}" );
print Dumper( \@items );

exit;

