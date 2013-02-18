
package SharePoint::CommandLine;

use strict;
use warnings;
use Data::Dumper;
use SharePoint::Connection;

use Getopt::Long;

sub get_options
{
	my( $class, $extra_opts, $extra_conf ) = @_;

	Getopt::Long::Configure("permute");

	my %opts = (
		show_help => 0,
		show_version => 0,
		debug => 0,
		credentials=>undef,
		site=>undef,
		%{$extra_opts},
	);
	my %optconf = (
		'help|?' => \$opts{show_help},
		'version' => \$opts{show_version},
		'debug' => \$opts{debug},
		'credentials-file:s' => \$opts{credentials},
		'site:s' => \$opts{site},
	);
	foreach my $k ( keys %{$extra_conf} )
	{
		$optconf{$k} = \$opts{$extra_conf->{$k}};
	}
	GetOptions( %optconf ) || show_usage();

	if( $opts{credentials} )
	{
		open( C, $opts{credentials} ) || 
			die( "Failed to read credentials file '".$opts{credentials}."': $!" );
		my $c = {};
		while( my $line = <C> )
		{
			chomp $line;
			$line =~ s/^\s*//;
			next if $line eq "";
			next if $line =~ m/^#/;
			if( !$line =~ m/:/ ) { die( "Bad credentials file '".$opts{credentials}."': Line without expected colon" ); }
			my($k,$v) = split( /\s*:\s*/, $line, 2 );
			$c->{$k} = $v;
		}
		if( !defined $c->{username} )
		{
			die( "Bad credentials file '".$opts{credentials}."': No username" ); 
		}
		if( !defined $c->{password} )
		{
			die( "Bad credentials file '".$opts{credentials}."': No password" ); 
		}
		$opts{username} = $c->{username};
		$opts{password} = $c->{password};
	}
	
	return %opts;
};

1;
