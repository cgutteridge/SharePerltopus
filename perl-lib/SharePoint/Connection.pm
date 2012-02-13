package SharePoint::Connection;

use LWP::UserAgent;
use LWP::Debug;
use Authen::NTLM;
use SOAP::Lite on_action => sub { "$_[0]$_[1]"; };
use Data::Dumper;

use strict;
use warnings;


# required opts
# -- endpoint
# -- username
# -- password
# optional
# -- debug (default: 0)
# -- on_error (default: ::die)

sub new
{
	my( $class, %opts ) = @_;

	my $self = bless {}, $class;
	$self->{opts} = \%opts;

	if( !defined $opts{debug} ) { $opts{debug} = 0; }
	#if( !defined $opts{on_error} ) { $opts{on_error} = &die; }

	if( $opts{debug} )
	{
		LWP::Debug::level('+');
		SOAP::Lite->import(+trace => 'all');
	}

	if( !defined $opts{endpoint} )
	{
		$opts{endpoint} = $opts{site}."/_vti_bin/lists.asmx";
	}

	$self->{soap} = SOAP::Lite->proxy( $opts{endpoint}, keep_alive => 1);
	$self->{soap}->uri("http://schemas.microsoft.com/sharepoint/soap/");

	# There's got to be a better way, but this does appear to work!
	eval "sub SOAP::Transport::HTTP::Client::get_basic_credentials { return ('$opts{username}' => '$opts{password}') };"; 
	
	return $self;
}

sub error
{
	my( $self, $msg ) = @_;

	if( !defined $self->{opts}->{on_error} ) { die $msg; }

	&{$self->{opts}->{on_error}}( $msg );
}
sub soapError 
{
	my( $self, $call ) = @_;

	my $msg = $call->faultstring().": ".$call->faultdetail()->{errorstring};
	return $self->error( $msg );
}


sub GetListCollection
{
	my( $self ) = @_;
	
	my $call = $self->{soap}->GetListCollection();
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetListCollectionResult/Lists/List') );
}

# nb. listName is a {234234} style ID!
sub GetList
{
	my( $self, $listName ) = @_;

	my $in_listName = SOAP::Data::name('listName' => $listName);

	my $call = $self->{soap}->GetList($in_listName);
	$self->soapError($call) if defined $call->fault();

	#return $self->attrFromList( $call->dataof('//GetListItemsResult/listitems/data/row') );
	my $r = {
		list => $call->dataof('//GetListResult/List')->attr, 
		fields => [ $self->attrFromList( $call->dataof( "//GetListResult/List/Fields/Field" ) )]
	};
	return $r;
	
}
# nb. listName is a {234234} style ID!
sub GetListItems
{
	my( $self, $listName, $viewName, $rowLimit ) = @_;

	$viewName = '' unless defined $viewName;
	$rowLimit = 99999 unless defined $rowLimit;

	my $in_listName = SOAP::Data::name('listName' => $listName);
	my $in_viewName = SOAP::Data::name('viewName' => $viewName);
	my $in_rowLimit = SOAP::Data::name('rowLimit' => $rowLimit);

	my $call = $self->{soap}->GetListItems($in_listName, $in_viewName, $in_rowLimit);
	$self->soapError($call) if defined $call->fault();

	return $self->attrFromList( $call->dataof('//GetListItemsResult/listitems/data/row') );
}

# It's kind of annoying to have to call attr on every list item by hand, so let's do it
# in a handy function (There may later turn out to be a reason to look elsewhere than in attr
# -- we'll see, I guess)
sub attrFromList
{
	my( $self, @list ) = @_;

	my @r = ();
	foreach my $item ( @list )
	{
		push @r, $item->attr;
	}
	return @r;
}

1;
