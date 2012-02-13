#!/usr/bin/perl -I/var/wwwsites/SharePerl/perl-lib


use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;

require "secrets.pl";
my $spc = sp_connect();
my $info =  $spc->GetList( "{58F8FB8E-04E3-4FC3-9ADB-DEE1B04BC346}" );
print Dumper( $info );

exit;

