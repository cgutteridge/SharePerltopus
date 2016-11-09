package SharePoint::Connection;

use LWP::UserAgent;
use LWP::Debug;
use Authen::NTLM;
use SOAP::Lite on_action => sub { "$_[0]$_[1]"; };
use Data::Dumper;
use MIME::Base64;
use Time::Local;

use strict;
use warnings;


# required opts
# -- endpoint
# -- username
# -- password
# -- site
# optional
# -- debug (default: 0)
# -- domain
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

	my $username = $opts{username};
	if( defined $opts{domain} )
	{
		$username = $opts{domain}."\\".$username;
	}

	# There's got to be a better way, but this does appear to work!
	delete $SOAP::Transport::HTTP::Client::{get_basic_credentials};
	eval "sub SOAP::Transport::HTTP::Client::get_basic_credentials { return ('$username' => '$opts{password}') };"; 
	
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
	my( $self, $listName, $viewName, $rowLimit, $where ) = @_;

	$viewName = '' unless defined $viewName;
	$rowLimit = 99999 unless defined $rowLimit;

	my $in_listName = SOAP::Data::name('listName' => $listName);
	my $in_viewName = SOAP::Data::name('viewName' => $viewName);
	my $in_rowLimit = SOAP::Data::name('rowLimit' => $rowLimit);
	my $in_query = undef;
	if( $where ) 
	{
		$in_query = SOAP::Data::name('query' => 
				\SOAP::Data->name("Query" => 
					\SOAP::Data->name("Where" =>
						\SOAP::Data->name( "dummy",  SOAP::Data->type( 'xml'=>$where ) ) ) ) );
	}

	my $call = $self->getListsEndpoint()->GetListItems($in_listName, $in_viewName, $in_query, $in_rowLimit);
	$self->soapError($call) if defined $call->fault();

	return $self->attrFromList( $call->dataof('//GetListItemsResult/listitems/data/row') );
}

# nb. listName is a {234234} style ID!
sub GetCalendarEvents
{
	my( $self, $listName, $viewName, $rowLimit, $where ) = @_;

	$viewName = '' unless defined $viewName;
	$rowLimit = 99999 unless defined $rowLimit;

	my $in_listName = SOAP::Data::name('listName' => $listName);
	my $in_viewName = SOAP::Data::name('viewName' => $viewName);
	my $in_rowLimit = SOAP::Data::name('rowLimit' => $rowLimit);
	my $in_query = SOAP::Data->type( 'xml'=>$where );

        my $overlap = "<DateRangesOverlap><FieldRef Name='EventDate' /><FieldRef Name='EndDate' /><FieldRef Name='RecurrenceID' /><Value Type='DateTime'><Year /></Value></DateRangesOverlap>";
	if( !$where ) 
	{
		$where = $overlap;
	}
	else
	{
		$where = "<And>$where$overlap</And>";
	}

	$in_query = SOAP::Data::name('query' => 
			\SOAP::Data->name("Query" => 
				\SOAP::Data->name("Where" =>
					\SOAP::Data->name( "dummy",  SOAP::Data->type( 'xml'=>$where ) ) ) ) );

	my $query_options = SOAP::Data::name('queryOptions' => 
				\SOAP::Data->name("QueryOptions" => 
					\SOAP::Data->name("ExpandRecurrence", SOAP::Data->type( 'xml'=>"<ExpandRecurrence>TRUE</ExpandRecurrence>" ))));

	my $call = $self->getListsEndpoint()->GetListItems($in_listName, $in_viewName, $in_query, $in_rowLimit, $query_options);
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

####### Higher level conversion functions

sub CalendarAsICAL
{
	my( $self, %opts ) = @_;

	my $map = {};
	my $listinfo = $self->GetList( $opts{list} );
	foreach my $field ( @{$listinfo->{fields}} )
	{
		$map->{$field->{DisplayName}} = "ows_".$field->{StaticName};
	}
	# always use real title
	$map->{"Title"}="ows_Title"; 
	if( defined $opts{view} && $opts{view} eq "" ) { delete $opts{view}; }

	my @items = $self->GetCalendarEvents( $opts{list}, $opts{view}, undef, $opts{where} );

	my @fields = ("Start Time", "End Time", "Title", "Description", "Unique Id");

	my $title = "Sharepoint Calendar";
	if( defined $opts{title} && $opts{title} ne "" ) { $title=$opts{title}; }

	my @output = ();

	push @output, "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SOTON//Sharepoint Calendar Fields//EN\r\nMETHOD:PUBLISH\r\nX-WR-CALNAME:$title\r\nX-WR-CALDESC:$title\r\nX-WR-TIMEZONE:Europe/London\r\n";

	foreach my $item ( @items )
	{
		push @output, "BEGIN:VEVENT\r\n";
		my @start = split(/[- :]/, $item->{$map->{'Start Time'}});
		my @end = split(/[- :]/, $item->{$map->{'End Time'}});
		# work out if this is a all day event, which sharepoint
		# exports as 000000 to 235900. If so make it an ics all day
		if( $start[3].$start[4].$start[5] eq "000000"
		 && $end[3].$end[4].$end[5] eq "235900" )
		{
			my $end_date_t = timelocal(0,0,0,$end[2],$end[1]-1,$end[0]-1900);
			$end_date_t += 24*60*60;
			my @end_date = localtime( $end_date_t );
			push @output, sprintf( "DTSTART;TZID=Europe/London;VALUE=DATE:%04d%02d%02d\r\n", @start );
			push @output, sprintf( "DTEND;TZID=Europe/London;VALUE=DATE:%04d%02d%02d\r\n", $end_date[5]+1900,$end_date[4]+1,$end_date[3]);
		}
		else
		{
			push @output, sprintf( "DTSTART;TZID=Europe/London:%04d%02d%02dT%02d%02d%02d\r\n", @start );
			push @output, sprintf( "DTEND;TZID=Europe/London:%04d%02d%02dT%02d%02d%02d\r\n", @end);
		}
		if ( !defined($item->{$map->{'Title'}}) ) { $item->{$map->{'Title'}} = "No Summary"; }
		push @output, "SUMMARY:".$item->{$map->{'Title'}}."\r\n";
		if ( !defined($item->{$map->{'Description'}}) ) { $item->{$map->{'Description'}} = "No Description"; }
		$item->{$map->{'Description'}} =~ s/[\r\n]/ /g;
		push @output, "DESCRIPTION:".$item->{$map->{'Description'}}."\r\n";
		if ( defined($item->{$map->{'Location'}}) ) { push @output, "LOCATION:".$item->{$map->{'Location'}}."\r\n"; }
		push @output, "END:VEVENT\r\n";
	}

	push @output, "END:VCALENDAR\r\n";

	return join( '', @output );
}

sub ListAsTSV
{
	my( $self, %opts ) = @_;

	my $map = {};
	my $listinfo = $self->GetList( $opts{list} );
	my $formats = {};
	foreach my $field ( @{$listinfo->{fields}} )
	{
		$map->{"ows_".$field->{StaticName}} = $field->{DisplayName};
		$formats->{"ows_".$field->{StaticName}} = $field->{Format} || "none";
	}
	if( defined $opts{view} && $opts{view} eq "" ) { delete $opts{view}; }

	my @items = $self->GetListItems( $opts{list}, $opts{view}, undef, $opts{where} );

	my @fields;
	if( defined $opts{'fields-list'} )
	{
		@fields  = @{ $opts{"fields-list"} };
	}
	if( !scalar @fields )
	{
		my $f={};
		foreach my $item ( @items )
		{
			CELL: foreach my $k ( sort keys %{$item} )
			{
				next CELL unless defined $map->{$k};
				$f->{$map->{$k}} = 1;
			}
		}
		@fields = sort keys %$f;
	}

	my @output = ();	
	my $row = join( "\t", @fields );
	$row =~ s/\n/\\n/g;
	push @output, $row."\n";
	foreach my $item ( @items )
	{
		my $cells = {};
		CELL: foreach my $k ( sort keys %{$item} )
		{
			next CELL unless defined $map->{$k};
			my $v = $item->{$k};
			my $name = $map->{$k};
			$v =~ s/^.*;#//;
			# url & email fields need more cleaning, but for now:
			# mailto:hs8@soton.ac.uk, hs8@soton.ac.uk
			if( $formats->{$k} eq "Hyperlink" )
			{
				$v=~s/^([^,]*),.*$/$1/;
			}
			# obviously tabs are a bad idea in a tsv value
			$v =~ s/\t/ /g;

			$cells->{$name} = $v;
		}
		my @values = ();
		foreach my $cell_name ( @fields )
		{
			my $v = $cells->{$cell_name};
			$v = "" unless defined $v;
			push @values, $v;
		}
		my $row = join( "\t", @values );
		$row =~ s/\n/\\n/g;
		push @output, $row."\n";
	}

	return join( '', @output );
}

1;
