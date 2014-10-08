#!/usr/bin/perl
use strict;
use Tk;
use Tk::DirTree;
use Cwd;
use DBI;
use Carp;
use open ':std', ':encoding(UTF-8)';

my $CSS = "
body
{
 font-family: Tahoma, Verdana, sans-serif;
 padding: 0;
 margin: 0;
 line-height:1.4;
 font-size: 12px;
 color: #000000;
}
h1
{
 background-color: #A0B0A0;
 margin: 0;
 padding: .25em;
 font-weight: bold;
 letter-spacing: 2px;
 font-size: 28px;
 text-align: center;
 border-bottom: 1px solid #303060;
}
h2
{
 margin-top: 20px;
 margin-left: 10px;
 padding: 0;
 border-bottom: 1px solid #C0C0C0;
 color: #002060;
 font-weight: normal;
 font-style: italic;
 letter-spacing: 2px;
}
table
{
 border: 0;
 padding: 0;
 margin: 0;
 border: 1px solid #303030;
}
th
{
 white-space: no-wrap;
 background-color: #A0B0A0;
 padding: 2px 20px;
 border-bottom: 3px solid #303060;
 border-right:1px solid #303060;
}
tr:nth-child(even)
{
 background-color:#D0E0D0;
}
tr:nth-child(odd)
{
 background-color:#D5EADF;
}
td
{
 padding:2px 6px;
 border-right:1px solid #303060;
 border-bottom:1px solid #303060;
}
td:hover
{
 background-color:#A0B0A0;
}
#notice
{
 font-style: italic;
 font-size: 12px;
 color: #BB9966;
 margin-left: 10px;
}
#header
{
 position: fixed;
 width: 100%;
 background-color: #FFFFFF;
 border-bottom: 3px solid #303060;
}
#content
{
 padding-top: 60px;
}
";
my $dbtype = "mssql";
my $mssql_host;
my $mssql_pass;
my $mssql_userid;
my $mssql_db;
my $sqlite_file;
my $mw;
my $cw;
my $con;
my $tablelist;
my $previewtext;
my $sql;
my $html_title;
my $title = "Database report";
my $html_author;
my $author = "";
my $filename = "export.html";
my $html_file;
my $folder = Cwd::cwd();
my $fp;
my $tablename;
my $csstext;
my $host = "localhost";
my $db = "master";
my $sid = "ORCL";
my $mydb = "mysql";
my $sqldb = "";
my $user = "";
my $pass = "";
my $driver = "";
my $dsn = "";
my $port = "3306";
my $instance = "SQLEXPRESS";
my $my_host;
my $my_userid;
my $my_pass;
my $my_db;
my $my_port;
my $oracle_sid;
my $oracle_userid;
my $oracle_pass;
my $oracle_host;
my $odbc_driver;
my $odbc_string;
my $sortby = "";
my $html_sortby;
my $html_images;
my $images = ".jpg .gif .png";
my $html_links;
my $links = "http:// https://";

# Show a message
sub msg
{
	my ($errmsg, $action) = @_;
	if($action == 1) { $mw->messageBox(-icon => 'error', -message => $errmsg, -title => 'Error', -type => 'Ok'); }
	elsif($action == 2) { $mw->messageBox(-icon => 'question', -message => $errmsg, -title => 'About', -type => 'Ok'); }
	else { $mw->messageBox(-icon => 'info', -message => $errmsg, -title => 'Info', -type => 'Ok'); }
}

# html options
sub set_options
{
	$author = $html_author->get;
	$title = $html_title->get;
	$filename = $html_file->get;
	$sortby = $html_sortby->get;
	$links = $html_links->get;
	$images = $html_images->get;
	$CSS = $csstext->get('1.0', 'end');
	$cw->destroy;
}

sub options
{
	$cw = $mw->Toplevel();
	$cw->title("Options");

	$cw->Label(-text => "HTML Title:")->grid(-column => 0, -row => 1, -padx => 10, -pady => 5);
	$html_title = $cw->Entry()->grid(-column => 1, -row => 1, -padx => 10, -pady => 5);
	$html_title->insert(0, $title);
	$cw->Label(-text => "Author name:")->grid(-column => 0, -row => 2, -padx => 10, -pady => 5);
	$html_author = $cw->Entry()->grid(-column => 1, -row => 2, -padx => 10, -pady => 5);
	$html_author->insert(0, $author);
	$cw->Label(-text => "Sort by:")->grid(-column => 0, -row => 3, -padx => 10, -pady => 5);
	$html_sortby = $cw->Entry()->grid(-column => 1, -row => 3, -padx => 10, -pady => 5);
	$html_sortby->insert(0, $sortby);
	$cw->Label(-text => "Links:")->grid(-column => 0, -row => 4, -padx => 10, -pady => 5);
	$html_links = $cw->Entry()->grid(-column => 1, -row => 4, -padx => 10, -pady => 5);
	$html_links->insert(0, $links);
	$cw->Label(-text => "Images:")->grid(-column => 0, -row => 5, -padx => 10, -pady => 5);
	$html_images = $cw->Entry()->grid(-column => 1, -row => 5, -padx => 10, -pady => 5);
	$html_images->insert(0, $images);
	my $html_folder = $cw->Scrolled('DirTree', -scrollbars => "se", -height => 15, -width => 40, -browsecmd => sub {$folder = shift}, -command => \&show_folder)->grid(-columnspan => 2, -row => 6, -padx => 5, -pady => 5);
	$html_folder->chdir($folder);
	$cw->Label(-text => "File name:")->grid(-column => 0, -row => 7, -padx => 10, -pady => 10);
	$html_file = $cw->Entry()->grid(-column => 1, -row => 7, -padx => 10, -pady => 10);
	$html_file->insert(0, $filename);
	$cw->Label(-text => "CSS:")->grid(-column => 2, -row => 0, -padx => 10, -pady => 5);
	$csstext = $cw->Scrolled("Text", -width => 50, -height => 30, -scrollbars => 'se')->grid(-row => 1, -rowspan => 7, -column => 2, -padx => 10, -pady => 5);
	$csstext->insert('end', $CSS);
	$cw->Button(-width => 15, -height => 1, -text => "Ok", -command => \&set_options)->grid(-columnspan => 3, -row => 8, -padx => 10, -pady => 10, -sticky => 'e'); 
	$cw->resizable(0, 0); 
	$cw->withdraw(); 
	$cw->Popup(); 
}

# make html file
sub make_html
{
	if(!$sql)
	{
		msg("Please select a table first.", 1);
		return;
	}
	my $path = File::Spec->catfile($folder, $filename);
	if(-e $path)
	{
		my $ok = $mw->messageBox(-icon => "question", -message => "File already exist. Overwrite?", -title => "Error", -type => "OkCancel");
		if($ok eq "Cancel") { return; }
	}
	eval { open($fp, ">$path") } or do { msg("Could not export to $path\n", 1); };
	if($fp)
	{
		print $fp "<html><head>\n";
		print $fp "<style type='text/css'>\n";
		print $fp $CSS;
		print $fp "</style>\n";
		if($title ne "") { print $fp "<title>" . $title . "</title>\n"; }
		print $fp "</head><body>\n";
		print $fp "<div id='header'><h1>" . $title . "</h1></div><div id='content'>\n";
		if($author ne "") { print $fp "<h2>by: " . $author . "</h2>\n"; }
		print $fp "<table tableborder=1><tr><th>\n";
		eval
		{
			if($sortby eq "") { $sql = $con->prepare("SELECT * FROM " . $tablename); }
			else { $sql = $con->prepare("SELECT * FROM " . $tablename . " ORDER BY " . $sortby); }
  			$sql->execute();
		} or do { msg("Could not access data.", 1); return; };
		my $names = $sql->{NAME};
		print $fp join('</th><th>', @$names);
		print $fp "</th></tr>\n";
		while (my @data = $sql->fetchrow_array())
		{
			print $fp "<tr>";
			foreach my $d (@data)
			{
				my $buf = "";
				foreach my $a (split(' ', $links))
				{
					if(index($d, $a) != -1) { $buf = "<a href='" . $d . "'>" . $d . "</a>"; }
				}
				foreach my $b (split(' ', $images))
				{
					if(index($d, $b) != -1) { $buf = "<img src='" . $d . "'>"; }
				}
				if($buf eq "") { $buf = $d; }
				print $fp "<td>" . $buf . "</td>";
			}
			print $fp "</tr>\n";
		}
		print $fp "</table></div>\n";
		print $fp "<p><span id='notice'>Data exported from " . $db . "." . $tablename . " [" . $dbtype . "] on " . localtime() . " using DB 2 HTML.</span></p>\n";
		print $fp "</body></html>\n";
		close($fp);
 		msg("Export done.", 0);
	}
}

# select table
sub select_table
{
	$tablename = $tablelist->get($tablelist->curselection());
	eval
	{
		$sql = $con->prepare("SELECT * FROM " . $tablename);
		$sql->execute();
	} or do { msg("Could not access data.", 1); };
	$previewtext->delete('1.0', 'end');
	my $names = $sql->{NAME};
	$previewtext->insert('end', join(', ', @$names) . "\n\n");
	while (my @data = $sql->fetchrow_array())
	{
		$previewtext->insert('end', join(', ', @data) . "\n");
	}
}

# connect to db
sub connect_db
{
	if($dbtype eq "mssql")
	{
		$host = $mssql_host->get();
		$db = $mssql_db->get();
		$user = $mssql_userid->get();
		$pass = $mssql_pass->get();
		eval
		{
			$con = DBI->connect("dbi:ODBC:Driver={SQL Server};Server=$host;DATABASE=$db;UID=$user;PWD=$pass");
		} or do { msg("Could not connect to database.\nPlease check your connection options.\n\n" . $DBI::errstr, 1); };
	}
	elsif($dbtype eq "odbc")
	{
		$db = $odbc_driver->get();
		$user = $odbc_string->get();
		eval
		{
			$con = DBI->connect("dbi:ODBC:Driver={" . $db . "};" . $user);
		} or do { msg("Could not connect to database.\nPlease check your connection options.\n\n" . $DBI::errstr, 1); };
	}
	elsif($dbtype eq "oracle")
	{
		$host = $oracle_host->get();
		$db = $oracle_sid->get();
		$user = $oracle_userid->get();
		$pass = $oracle_pass->get();
		eval
		{
			$con = DBI->connect("dbi:Oracle://$host/$db", $user, $pass);
		} or do { msg("Could not connect to database.\nPlease check your connection options.\n\n" . $DBI::errstr, 1); };
	}
	elsif($dbtype eq "mysql")
	{
		$db = $my_db->get();
		$user = $my_userid->get();
		$pass = $my_pass->get();
		$host = $my_host->get();
		$port = $my_port->get();
		eval
		{
			$con = DBI->connect("dbi:mysql:database=$db;host=$host;port=$port", $user, $pass);
		} or do { msg("Could not connect to database.\nPlease check your connection options.\n\n" . $DBI::errstr, 1); };
	}
	else
	{
		$db = $sqlite_file->get();
		eval
		{
			if(!-e $db) { croak("SQLite: File doesn't exist."); }
			$con = DBI->connect("dbi:SQLite:dbname=$db");
		} or do { msg("Could not connect to database.\nPlease check your connection options.\n\n" . $@, 1); };
	}
	if($con)
	{
		$cw->destroy;

		$mw->Label(-font => [-weight => 'bold'], -text => "Tables")->grid(-row => 0, -padx => 10, -pady => 5);
		$tablelist = $mw->Scrolled("Listbox", -width => 30, -height => 30, -selectmode => "single", -scrollbars => 'se')->grid(-row => 1, -padx => 10, -pady => 5);
		$mw->Button(-width => 15, -height => 1, -text => "Select", -command => \&select_table)->grid(-row => 2, -padx => 10, -pady => 5);
		$mw->Label(-font => [-weight => 'bold'], -text => "Data preview")->grid(-row => 0, -column => 1, -padx => 10, -columnspan => 3, -pady => 5);
		$previewtext = $mw->Scrolled("Text", -width => 60, -height => 30, -scrollbars => 'se')->grid(-row => 1, -column => 1, -columnspan => 3, -padx => 10, -pady => 5);
		$mw->Button(-width => 15, -height => 1, -text => "Options", -command => \&options)->grid(-row => 2, -column => 1, -padx => 10, -pady => 5);
		$mw->Button(-width => 15, -height => 1, -text => "Export", -command => \&make_html)->grid(-row => 2, -column => 2, -padx => 10, -pady => 5);
		$mw->Button(-width => 15, -height => 1, -text => "About", -command => sub { msg("DB 2 HTML v1.0 - by Patrick Lambert [http://dendory.net]\n\nThis utility allows you to export data from a database into an HTML file using custom options.\n\nUsage:\n\tdb2html -dbtype <mssql|oracle|mysql|sqlite|odbc> -host <hostname> -instance <instance> -port <port> -user <user id> -pass <password> -db <database> -sid <SID> -driver <ODBC driver> -dsn <ODBC DSN> -title <HTML title> -author <author> -sortby <header> -links <links> -images <files> -filename <export file> -folder <export folder>", 2); })->grid(-row => 2, -column => 3, -padx => 10, -pady => 5);

		my @tables = $con->tables();
		$tablelist->insert('end', @tables);

		$mw->resizable(0, 0); 
		$mw->withdraw(); 
		$mw->Popup(); 
	}
}

# cmd line options
while($#ARGV > -1)
{
	if($ARGV[0] eq "-css")
	{
		shift(@ARGV);
		$CSS = $ARGV[0];
	}
	elsif($ARGV[0] eq "-dbtype")
	{
		shift(@ARGV);
		$dbtype = $ARGV[0];
	}
	elsif($ARGV[0] eq "-db")
	{
		shift(@ARGV);
		$db = $ARGV[0];
		$mydb = $ARGV[0];
		$sqldb = $ARGV[0];
	}
	elsif($ARGV[0] eq "-user")
	{
		shift(@ARGV);
		$user = $ARGV[0];
	}
	elsif($ARGV[0] eq "-pass")
	{
		shift(@ARGV);
		$pass = $ARGV[0];
	}
	elsif($ARGV[0] eq "-filename")
	{
		shift(@ARGV);
		$filename = $ARGV[0];
	}
	elsif($ARGV[0] eq "-folder")
	{
		shift(@ARGV);
		$folder = $ARGV[0];
	}
	elsif($ARGV[0] eq "-sortby")
	{
		shift(@ARGV);
		$sortby = $ARGV[0];
	}
	elsif($ARGV[0] eq "-port")
	{
		shift(@ARGV);
		$port = $ARGV[0];
	}
	elsif($ARGV[0] eq "-host")
	{
		shift(@ARGV);
		$host = $ARGV[0];
	}
	elsif($ARGV[0] eq "-instance")
	{
		shift(@ARGV);
		$instance = $ARGV[0];
	}
	elsif($ARGV[0] eq "-author")
	{
		shift(@ARGV);
		$author = $ARGV[0];
	}
	elsif($ARGV[0] eq "-title")
	{
		shift(@ARGV);
		$title = $ARGV[0];
	}
	elsif($ARGV[0] eq "-images")
	{
		shift(@ARGV);
		$images = $ARGV[0];
	}
	elsif($ARGV[0] eq "-links")
	{
		shift(@ARGV);
		$links = $ARGV[0];
	}
	elsif($ARGV[0] eq "-driver")
	{
		shift(@ARGV);
		$driver = $ARGV[0];
	}
	elsif($ARGV[0] eq "-dsn")
	{
		shift(@ARGV);
		$dsn = $ARGV[0];
	}
	elsif($ARGV[0] eq "-sid")
	{
		shift(@ARGV);
		$sid = $ARGV[0];
	}
	else
	{
		print STDERR "Unknown option: " . $ARGV[0] . "\n";
	}
	shift(@ARGV);
}

# make UI
$mw = MainWindow->new;
$mw->title("DB 2 HTML");

$cw = $mw->Toplevel();
$cw->title("Connection");

$cw->Label(-font => [-weight => 'bold'], -text => "Database connection")->grid(-columnspan => 4, -row => 0, -padx => 10, -pady => 10);
$cw->Radiobutton(-text => "Microsoft SQL Server", -value => "mssql", -variable => \$dbtype)->grid(-columnspan => 2, -row => 1, -padx => 10, -pady => 5, -sticky => 'w');
$cw->Label(-text => "Host \\ Instance:")->grid(-column => 0, -row => 2, -padx => 10, -pady => 5);
$mssql_host = $cw->Entry()->grid(-column => 1, -row => 2, -padx => 10, -pady => 5);
$mssql_host->insert(0, $host . "\\" . $instance);
$cw->Label(-text => "User ID:")->grid(-column => 0, -row => 3, -padx => 10, -pady => 5);
$mssql_userid = $cw->Entry()->grid(-column => 1, -row => 3, -padx => 10, -pady => 5);
$mssql_userid->insert(0, $user);
$cw->Label(-text => "Password:")->grid(-column => 0, -row => 4, -padx => 10, -pady => 5);
$mssql_pass = $cw->Entry(-show => '*')->grid(-column => 1, -row => 4, -padx => 10, -pady => 5);
$mssql_pass->insert(0, $pass);
$cw->Label(-text => "Database:")->grid(-column => 0, -row => 5, -padx => 10, -pady => 5);
$mssql_db = $cw->Entry()->grid(-column => 1, -row => 5, -padx => 10, -pady => 5);
$mssql_db->insert(0, $db);
$cw->Radiobutton(-text => "Oracle Server", -value => "oracle", -variable => \$dbtype)->grid(-columnspan => 2, -row => 6, -padx => 10, -pady => 5, -sticky => 'w');
$cw->Label(-text => "Hostname:")->grid(-column => 0, -row => 7, -padx => 10, -pady => 5);
$oracle_host = $cw->Entry()->grid(-column => 1, -row => 7, -padx => 10, -pady => 5);
$oracle_host->insert(0, $host);
$cw->Label(-text => "SID:")->grid(-column => 0, -row => 8, -padx => 10, -pady => 5);
$oracle_sid = $cw->Entry()->grid(-column => 1, -row => 8, -padx => 10, -pady => 5);
$oracle_sid->insert(0, $sid);
$cw->Label(-text => "User ID:")->grid(-column => 0, -row => 9, -padx => 10, -pady => 5);
$oracle_userid = $cw->Entry()->grid(-column => 1, -row => 9, -padx => 10, -pady => 5);
$oracle_userid->insert(0, $user);
$cw->Label(-text => "Password:")->grid(-column => 0, -row => 10, -padx => 10, -pady => 5);
$oracle_pass = $cw->Entry(-show => '*')->grid(-column => 1, -row => 10, -padx => 10, -pady => 5);
$oracle_pass->insert(0, $pass);
$cw->Radiobutton(-text => "MySQL", -value => "mysql", -variable => \$dbtype)->grid(-columnspan => 2, -row => 1, -column => 2, -padx => 10, -pady => 5, -sticky => 'w');
$cw->Label(-text => "Hostname:")->grid(-column => 2, -row => 2, -padx => 10, -pady => 5);
$my_host = $cw->Entry()->grid(-column => 3, -row => 2, -padx => 10, -pady => 5);
$my_host->insert(0, $host);
$cw->Label(-text => "Port:")->grid(-column => 2, -row => 3, -padx => 10, -pady => 5);
$my_port = $cw->Entry()->grid(-column => 3, -row => 3, -padx => 10, -pady => 5);
$my_port->insert(0, $port);
$cw->Label(-text => "User ID:")->grid(-column => 2, -row => 4, -padx => 10, -pady => 5);
$my_userid = $cw->Entry()->grid(-column => 3, -row => 4, -padx => 10, -pady => 5);
$my_userid->insert(0, $user);
$cw->Label(-text => "Password:")->grid(-column => 2, -row => 5, -padx => 10, -pady => 5);
$my_pass = $cw->Entry(-show => '*')->grid(-column => 3, -row => 5, -padx => 10, -pady => 5);
$my_pass->insert(0, $pass);
$cw->Label(-text => "Database:")->grid(-column => 2, -row => 6, -padx => 10, -pady => 5);
$my_db = $cw->Entry()->grid(-column => 3, -row => 6, -padx => 10, -pady => 5);
$my_db->insert(0, $mydb);
$cw->Radiobutton(-text => "SQLite 3", -value => "sqlite", -variable => \$dbtype)->grid(-column => 2, -columnspan => 2, -row => 7, -padx => 10, -pady => 5, -sticky => 'w');
$cw->Label(-text => "File name:")->grid(-column => 2, -row => 8, -padx => 10, -pady => 5);
$sqlite_file = $cw->Entry()->grid(-column => 3, -row => 8, -padx => 10, -pady => 5);
$sqlite_file->insert(0, $sqldb);
$cw->Radiobutton(-text => "Generic ODBC", -value => "odbc", -variable => \$dbtype)->grid(-column => 2, -columnspan => 2, -row => 9, -padx => 10, -pady => 5, -sticky => 'w');
$cw->Label(-text => "Driver name:")->grid(-column => 2, -row => 10, -padx => 10, -pady => 5);
$odbc_driver = $cw->Entry()->grid(-column => 3, -row => 10, -padx => 10, -pady => 5);
$odbc_driver->insert(0, $driver);
$cw->Label(-text => "DSN:")->grid(-column => 2, -row => 11, -padx => 10, -pady => 5);
$odbc_string = $cw->Entry()->grid(-column => 3, -row => 11, -padx => 10, -pady => 5);
$odbc_string->insert(0, $dsn);
$cw->Button(-width => 15, -height => 1, -text => "Connect", -command => \&connect_db)->grid(-columnspan => 4, -column => 0, -row => 12, -padx => 10, -pady => 10); 

$mw->withdraw();  # hide main window for now
$cw->protocol('WM_DELETE_WINDOW', sub { exit });  # closing connection window exits the app

$cw->resizable(0, 0); 
$cw->withdraw(); 
$cw->Popup(); 

# main loop
MainLoop;