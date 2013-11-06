package SharePoint::Connection;

use LWP::UserAgent;
use LWP::Debug;
use Authen::NTLM;
use SOAP::Lite on_action => sub { "$_[0]$_[1]"; };
use Data::Dumper;
use MIME::Base64;

use strict;
use warnings;


# required opts
# -- endpoint
# -- username
# -- password
# -- site
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

	# There's got to be a better way, but this does appear to work!
	eval "sub SOAP::Transport::HTTP::Client::get_basic_credentials { return ('$opts{username}' => '$opts{password}') };"; 
	
	return $self;
}

sub getUserGroupEndpoint
{
	my( $self ) = @_;

	if( !defined $self->{soap}->{usergroup} )
	{
		my $endpoint = $self->{opts}->{site}."/_vti_bin/usergroup.asmx";
		$self->{soap}->{usergroup} = SOAP::Lite->proxy( $endpoint, keep_alive => 1);
		$self->{soap}->{usergroup}->uri("http://schemas.microsoft.com/sharepoint/soap/directory/");
	}

	return $self->{soap}->{usergroup};
}

sub getListsEndpoint
{
	my( $self ) = @_;

	if( !defined $self->{soap}->{lists} )
	{
		my $endpoint = $self->{opts}->{site}."/_vti_bin/lists.asmx";
		$self->{soap}->{lists} = SOAP::Lite->proxy( $endpoint, keep_alive => 1);
		$self->{soap}->{lists}->uri("http://schemas.microsoft.com/sharepoint/soap/");
	}

	return $self->{soap}->{lists};
}

sub getCopyEndpoint
{
	my( $self ) = @_;

	if( !defined $self->{soap}->{copy} )
	{
		my $endpoint = $self->{opts}->{site}."/_vti_bin/copy.asmx";
		$self->{soap}->{copy} = SOAP::Lite->proxy( $endpoint, keep_alive => 1);
		$self->{soap}->{copy}->uri("http://schemas.microsoft.com/sharepoint/soap/");
	}

	return $self->{soap}->{copy};
}

sub getWebsEndpoint
{
	my( $self ) = @_;

	if( !defined $self->{soap}->{webs} )
	{
		my $endpoint = $self->{opts}->{site}."/_vti_bin/webs.asmx";
		$self->{soap}->{webs} = SOAP::Lite->proxy( $endpoint, keep_alive => 1);
		$self->{soap}->{webs}->uri("http://schemas.microsoft.com/sharepoint/soap/");
	}

	return $self->{soap}->{webs};
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

	my $msg = $call->faultstring().": ";
	if( $call->faultdetail() ne "" )
	{
		$msg.=$call->faultdetail()->{errorstring};
	}
	return $self->error( $msg );
}

sub GetItem
{
	my( $self, $Url ) = @_;
	
	my $in_Url = SOAP::Data::name('Url' => $Url );

	my $call = $self->getCopyEndpoint()->GetItem($in_Url);
	$self->soapError($call) if defined $call->fault();

	return MIME::Base64::decode( $call->dataof('//GetItemResponse/Stream' )->value() );
}

sub GetAttachmentCollection
{
	my( $self, $listName, $listItemID ) = @_;
	
	my $in_listName = SOAP::Data::name('listName' => $listName );
	my $in_listItemID = SOAP::Data::name( 'listItemID' => $listItemID )->type( 'string' );

	my $call = $self->getListsEndpoint()->GetAttachmentCollection($in_listName, $in_listItemID);
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetAttachmentCollectionResponse' ) );
}

sub GetWebCollection
{
	my( $self ) = @_;
	
	my $call = $self->getWebsEndpoint()->GetWebCollection();
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetWebCollectionResult/Webs/Web') );
}

sub GetListCollection
{
	my( $self ) = @_;
	
	my $call = $self->getListsEndpoint()->GetListCollection();
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetListCollectionResult/Lists/List') );
}

# nb. listName is a {234234} style ID!
sub GetList
{
	my( $self, $listName ) = @_;

	my $in_listName = SOAP::Data::name('listName' => $listName);

	my $call = $self->getListsEndpoint()->GetList($in_listName);
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

	my $call = $self->getListsEndpoint()->GetListItems($in_listName, $in_viewName, $in_rowLimit);
	$self->soapError($call) if defined $call->fault();

	return $self->attrFromList( $call->dataof('//GetListItemsResult/listitems/data/row') );
}

# nb. listName is a {234234} style ID!
sub GetCalendarEvents
{
	my( $self, $listName, $viewName, $rowLimit ) = @_;

	$viewName = '' unless defined $viewName;
	$rowLimit = 99999 unless defined $rowLimit;

	my $in_listName = SOAP::Data::name('listName' => $listName);
	my $in_viewName = SOAP::Data::name('viewName' => $viewName);
	my $in_rowLimit = SOAP::Data::name('rowLimit' => $rowLimit);
	my $query_options = SOAP::Data::name('queryOptions' => \SOAP::Data->name("QueryOptions" => \SOAP::Data->name("ExpandRecurrence", "TRUE")));

	my $call = $self->getListsEndpoint()->GetListItems($in_listName, $in_viewName, $in_rowLimit, $query_options);
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

######### Groups

sub GetGroupCollectionFromSite
{
	my( $self ) = @_;
	
	my $call = $self->getUserGroupEndpoint()->GetGroupCollectionFromSite();
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetGroupCollectionFromSite/Groups/Group') );
}

sub GetUserCollectionFromGroup
{
	my( $self, $groupName ) = @_;
	
	my $in_groupName = SOAP::Data::name( 'groupName' => $groupName );
	my $call = $self->getUserGroupEndpoint()->GetUserCollectionFromGroup( $in_groupName );
	$self->soapError($call) if defined $call->fault();
	
	return $self->attrFromList( $call->dataof('//GetUserCollectionFromGroup/Users/User') );
}


1;
